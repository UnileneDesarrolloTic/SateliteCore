using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Comercial
{
    public class ReporteOrdenFabricacionCaja
    {
        public string GenerarReporteCaja(IEnumerable<FormatoEstructuraObtenerOrdenFabricacion> dato)
        {

            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Orden Fabricacion Caja");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);
                TextoNegrita(worksheet);


                worksheet.Cells["A2"].Value = "Orde Fabricación";
                worksheet.Cells["A2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A2"].Style.Font.Size = 12;
                worksheet.Cells["A2"].Style.WrapText = true;
                worksheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["B2"].Value = "F.Producción";
                worksheet.Cells["B2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B2"].Style.Font.Size = 12;
                worksheet.Cells["B2"].Style.WrapText = true;
                worksheet.Cells["B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["C2"].Value = "Item";
                worksheet.Cells["C2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C2"].Style.Font.Size = 12;
                worksheet.Cells["C2"].Style.WrapText = true;
                worksheet.Cells["C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D2"].Value = "Código";
                worksheet.Cells["D2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D2"].Style.Font.Size = 12;
                worksheet.Cells["D2"].Style.WrapText = true;
                worksheet.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E2"].Value = "Descripción Local";
                worksheet.Cells["E2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E2"].Style.Font.Size = 12;
                worksheet.Cells["E2"].Style.WrapText = true;
                worksheet.Cells["E2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F2"].Value = "Número Lote";
                worksheet.Cells["F2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F2"].Style.Font.Size = 12;
                worksheet.Cells["F2"].Style.WrapText = true;
                worksheet.Cells["F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G2"].Value = "Cliente";
                worksheet.Cells["G2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G2"].Style.Font.Size = 12;
                worksheet.Cells["G2"].Style.WrapText = true;
                worksheet.Cells["G2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                worksheet.Cells["H2"].Value = "Marca";
                worksheet.Cells["H2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H2"].Style.Font.Size = 12;
                worksheet.Cells["H2"].Style.WrapText = true;
                worksheet.Cells["H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["I2"].Value = "Contra Muestra";
                worksheet.Cells["I2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I2"].Style.Font.Size = 12;
                worksheet.Cells["I2"].Style.WrapText = true;
                worksheet.Cells["I2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J2"].Value = "Numero Caja";
                worksheet.Cells["J2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J2"].Style.Font.Size = 12;
                worksheet.Cells["J2"].Style.WrapText = true;
                worksheet.Cells["J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["K2"].Value = "Fecha Ingreso";
                worksheet.Cells["K2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K2"].Style.Font.Size = 12;
                worksheet.Cells["K2"].Style.WrapText = true;
                worksheet.Cells["K2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 3;

                foreach (FormatoEstructuraObtenerOrdenFabricacion item in dato)
                {
                    worksheet.Row(row).Height = 25.5;

                    worksheet.Cells["A" + row].Value = item.OrdenFabricacion;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["B" + row].Value = item.FechaProduccion;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                    worksheet.Cells["C" + row].Value = item.Item;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["D" + row].Value = item.NumeroParte;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["E" + row].Value = item.DescripcionLocal;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["F" + row].Value = item.Lote;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                    worksheet.Cells["G" + row].Value = item.Cliente;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["H" + row].Value = item.Marca;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["I" + row].Value = item.ContraMuestra;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["J" + row].Value = item.NumeroCaja;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + row].Style.Font.Size = 10;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["K" + row].Value = item.FechaIngreso;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["K" + row].Style.Font.Size = 10;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "dd/MM/yyyy";


                    row++;
                }



                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

            }

            return reporte;
        }

        private static void AlineacionesTexto(ExcelWorksheet worksheet)
        {


        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 16.86 + 2.71;
            worksheet.Column(2).Width = 12.57 + 2.71;
            worksheet.Column(3).Width = 13.71 + 2.71;
            worksheet.Column(4).Width = 20.00 + 2.71;
            worksheet.Column(5).Width = 75.86 + 2.71;
            worksheet.Column(6).Width = 10.00 + 2.71;
            worksheet.Column(7).Width = 40.71 + 2.71;
            worksheet.Column(8).Width = 7.00 + 2.71;
            worksheet.Column(9).Width = 10.00 + 2.71;
            worksheet.Column(10).Width = 10.40 + 2.71;
            worksheet.Column(11).Width = 10.40 + 2.71;
            worksheet.Column(12).Width = 10.40 + 2.71;

        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1:K1"].Merge = true;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A2:K2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D8D8D8"));
        }

        
        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
