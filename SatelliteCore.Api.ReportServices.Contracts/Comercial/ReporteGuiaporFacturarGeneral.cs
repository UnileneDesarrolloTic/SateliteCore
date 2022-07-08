using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Comercial
{
    public class ReporteGuiaporFacturarGeneral
    {
        public string ExportarListarGuiaPorFacturaGeneral(IEnumerable<FormatoGuiaPorFacturarGeneralModel> ListaGuiaPorFacturaGeneral)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Guia No Facturar");

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                PintarCeldas(worksheet);

                worksheet.Cells["A1"].Value = "CHECK";
                worksheet.Cells["A1"].Style.Font.Size = 10;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["B1"].Value = "SERIE";
                worksheet.Cells["B1"].Style.Font.Size = 10;
                worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["C1"].Value = "NUMERO";
                worksheet.Cells["C1"].Style.Font.Size = 10;
                worksheet.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D1"].Value = "CLIENTE";
                worksheet.Cells["D1"].Style.Font.Size = 10;
                worksheet.Cells["D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["E1"].Value = "FACTURA";
                worksheet.Cells["E1"].Style.Font.Size = 10;
                worksheet.Cells["E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["F1"].Value = "F.FACTURA";
                worksheet.Cells["F1"].Style.Font.Size = 10;
                worksheet.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["G1"].Value = "E.FACTURA";
                worksheet.Cells["G1"].Style.Font.Size = 10;
                worksheet.Cells["G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["H1"].Value = "F.DOCUMENTO";
                worksheet.Cells["H1"].Style.Font.Size = 10;
                worksheet.Cells["H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["I1"].Value = "NUMERO_ORDEN";
                worksheet.Cells["I1"].Style.Font.Size = 10;
                worksheet.Cells["I1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["J1"].Value = "DESCRIPCION";
                worksheet.Cells["J1"].Style.Font.Size = 10;
                worksheet.Cells["J1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["K1"].Value = "CANTIDAD";
                worksheet.Cells["K1"].Style.Font.Size = 10;
                worksheet.Cells["K1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["L1"].Value = "NUMERO PROCESO";
                worksheet.Cells["L1"].Style.Font.Size = 10;
                worksheet.Cells["L1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["M1"].Value = "P.PARTIDA";
                worksheet.Cells["M1"].Style.Font.Size = 10;
                worksheet.Cells["M1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["N1"].Value = "ESTADO";
                worksheet.Cells["N1"].Style.Font.Size = 10;
                worksheet.Cells["N1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int row = 2;

                foreach (FormatoGuiaPorFacturarGeneralModel item in ListaGuiaPorFacturaGeneral)
                {
                    worksheet.Row(row).Height = 25.5;

                    worksheet.Cells["A" + row].Value = item.comentariosEntrega ? "SI": "NO";
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["B" + row].Value = item.serienumero;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["C" + row].Value = item.guianumero;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["D" + row].Value = item.Cliente;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["E" + row].Value = item.FacturaNumero;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["F" + row].Value = item.FacturaFecha;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + row].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["G" + row].Value = item.EstadoFacturacion;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["H" + row].Value = item.FechaDocumento;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["I" + row].Value = item.ReferenciaNumeroOrden;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["J" + row].Value = item.Descripcion;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + row].Style.Font.Size = 10;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["K" + row].Value = item.Cantidad;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["K" + row].Style.Font.Size = 10;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";

                    worksheet.Cells["L" + row].Value = item.LicitacionNumeroProceso;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["L" + row].Style.Font.Size = 10;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["M" + row].Value = item.ReprogramacionPuntoPartida;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["M" + row].Style.Font.Size = 10;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["N" + row].Value = item.estado;
                    worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["N" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["N" + row].Style.Font.Size = 10;
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    row++ ;

                }

                


                    file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }


            return reporte;
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 4.43 + 2.71;
            worksheet.Column(2).Width = 8.43 + 2.71;
            worksheet.Column(3).Width = 10.43 + 2.71;
            worksheet.Column(4).Width = 45.43 + 2.71;
            worksheet.Column(5).Width = 10.43 + 2.71;
            worksheet.Column(6).Width = 10.43 + 2.71;
            worksheet.Column(7).Width = 11.86 + 2.71;
            worksheet.Column(8).Width = 13.43 + 2.71;
            worksheet.Column(9).Width = 13.29 + 2.71;
            worksheet.Column(10).Width = 40 + 2.71;
            worksheet.Column(11).Width = 8.43 + 2.71;
            worksheet.Column(12).Width = 17.43 + 2.71;
            worksheet.Column(13).Width = 8.43 + 2.71;
            worksheet.Column(14).Width = 8.43 + 2.71;
            worksheet.Column(15).Width = 10.43 + 2.71;

        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {
            // worksheet.Cells["N14,J5"].Style.Font.UnderLine = true;
            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D8D8D8"));

        }
    }  
}
