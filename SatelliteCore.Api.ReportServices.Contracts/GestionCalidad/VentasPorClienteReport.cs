using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.GestionCalidad
{
    public class VentasPorClienteReport
    {
        public static string Exportar(Image logoUnilene, List<VentasPorClienteDTO> ventas)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Ventas por cliente");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(182, 60);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["J2"].Value = "REPORTE DE VENTAS POR CLIENTE";
                workSheet.Cells["J2"].Style.Font.Name = "Arial";
                workSheet.Cells["J2"].Style.Font.Size = 18;
                workSheet.Cells["J2"].Style.Font.Bold = true;

                workSheet.Cells["B5"].Value = "Tipo Doc";
                workSheet.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["C5"].Value = "Nro Doc";
                workSheet.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D5"].Value = "Fecha Doc";
                workSheet.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E5"].Value = "Estado";
                workSheet.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F5"].Value = "Tipo Venta";
                workSheet.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G5"].Value = "Pedido";
                workSheet.Cells["G5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H5"].Value = "Item";
                workSheet.Cells["H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I5"].Value = "N° Parte";
                workSheet.Cells["I5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["J5"].Value = "Linea";
                workSheet.Cells["J5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["K5"].Value = "Familia";
                workSheet.Cells["K5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["L5"].Value = "Sub Familia";
                workSheet.Cells["L5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["M5"].Value = "Lote";
                workSheet.Cells["M5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["N5"].Value = "Item Serie";
                workSheet.Cells["N5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["O5"].Value = "Descripción";
                workSheet.Cells["O5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["P5"].Value = "Und";
                workSheet.Cells["P5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["Q5"].Value = "Entregado";
                workSheet.Cells["Q5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["R5"].Value = "Precio Und";
                workSheet.Cells["R5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["S5"].Value = "Monto"; 
                workSheet.Cells["S5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["T5"].Value = "Total";
                workSheet.Cells["T5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["U5"].Value = "Comentario";
                workSheet.Cells["U5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);




                int index = 6;

                foreach (VentasPorClienteDTO venta in ventas)
                {
                    workSheet.Cells["B" + index].Value = venta.TipoDocumento;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = venta.NumeroDocumento;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = venta.FechaDocumento;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = venta.Estado;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = venta.TipoVenta;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = venta.ComercialPedidoNumero;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = venta.ItemCodigo;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = venta.NumeroDeParte;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = venta.Linea;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = venta.Familia;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + index].Value = venta.SubFamilia;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + index].Value = venta.Lote;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["N" + index].Value = venta.ItemSerie;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["O" + index].Value = venta.Descripcion;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + index].Value = venta.UnidadCodigo;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["Q" + index].Value = venta.CantidadEntregada;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["R" + index].Value = venta.PrecioUnitario;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["S" + index].Value = venta.Monto;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["T" + index].Value = venta.MontoTotal;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["U" + index].Value = venta.Comentarios;
                    workSheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

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

            workSheet.Cells["B5:U5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["B6:U"+ index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            workSheet.Cells["B5:U5"].Style.Font.Size = 12;
            workSheet.Cells["B5:U5"].Style.Font.Bold = true;
            workSheet.Cells["B6:U" + index].Style.WrapText = true;

            workSheet.Cells["D6:D" + index].Style.Numberformat.Format = "dd/MM/yyyy hh:mm";
            workSheet.Cells["R6:T" + index].Style.Numberformat.Format = "#,##0.00";
            workSheet.Cells["Q6:Q" + index].Style.Numberformat.Format = "#,##0";

            workSheet.Row(5).Style.Font.Size = 12;

            workSheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Column(16).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            workSheet.Column(3).Width = 14;
            workSheet.Column(4).Width = 17;
            workSheet.Column(7).Width = 14;
            workSheet.Column(8).Width = 13;
            workSheet.Column(9).Width = 26;
            workSheet.Column(10).Width = 23;
            workSheet.Column(11).Width = 27;
            workSheet.Column(12).Width = 18;
            workSheet.Column(13).Width = 17;
            workSheet.Column(14).Width = 11;
            workSheet.Column(15).Width = 65;
            workSheet.Column(17).Width = 11;
            workSheet.Column(18).Width = 11;
            workSheet.Column(19).Width = 11;
            workSheet.Column(20).Width = 11;
            workSheet.Column(21).Width = 70;

            workSheet.Cells["B5:U5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }

    }
}
