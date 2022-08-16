using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Logistica
{
    public class ReporteItemVentas
    {
        public string GenerarReporte(IEnumerable<DatosFormatoItemVentas> datos)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Item Ventas");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);
                TextoNegrita(worksheet);



                worksheet.Cells["A3"].Value = "Item";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3"].Style.Font.Size = 10;
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B3"].Value = "Descripción Local";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B3"].Style.Font.Size = 10;
                worksheet.Cells["B3"].Style.WrapText = true;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C3"].Value = "UnidadMedida";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C3"].Style.Font.Size = 10;
                worksheet.Cells["C3"].Style.WrapText = true;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D3"].Value = "NumeroDeParte";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D3"].Style.Font.Size = 10;
                worksheet.Cells["D3"].Style.WrapText = true;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E3"].Value = "SubFamilia";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E3"].Style.Font.Size = 10;
                worksheet.Cells["E3"].Style.WrapText = true;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F3"].Value = "SubFamilia";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F3"].Style.Font.Size = 10;
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F3"].Value = "Nombre Marca";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F3"].Style.Font.Size = 10;
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G3"].Value = "Presentacion";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G3"].Style.Font.Size = 10;
                worksheet.Cells["G3"].Style.WrapText = true;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["H3"].Value = "StockActual";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H3"].Style.Font.Size = 10;
                worksheet.Cells["H3"].Style.WrapText = true;
                worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["I3"].Value = "StockComprometido";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I3"].Style.Font.Size = 9;
                worksheet.Cells["I3"].Style.WrapText = true;
                worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["J3"].Value = "StockDisponible";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J3"].Style.Font.Size = 10;
                worksheet.Cells["J3"].Style.WrapText = true;
                worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //detalle 
                int row = 4;

                foreach (DatosFormatoItemVentas rowitem in datos)
                {
                    worksheet.Row(row).Height = 14.25;

                    worksheet.Cells["A" + row].Value = rowitem.Item;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["B" + row].Value = rowitem.DescripcionLocal;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["C" + row].Value = rowitem.UnidadCodigo;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["D" + row].Value = rowitem.NumeroDeParte;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["E" + row].Value = rowitem.SubFamilia;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["F" + row].Value = rowitem.NombreMarca;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["G" + row].Value = rowitem.Presentacion;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["H" + row].Value = rowitem.StockActual;
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["I" + row].Value = rowitem.StockComprometido;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["I" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["J" + row].Value = rowitem.StockDisponible;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["J" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + row].Style.Font.Size = 10;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


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
        {


        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {

            worksheet.Column(1).Width = 20.86 + 2.71;
            worksheet.Column(2).Width = 80.43 + 2.71;
            worksheet.Column(3).Width = 14.71 + 2.71;
            worksheet.Column(4).Width = 33.14 + 2.71;
            worksheet.Column(5).Width = 17.57 + 2.71;
            worksheet.Column(6).Width = 20.57 + 2.71;
            worksheet.Column(7).Width = 20.57 + 2.71;
            worksheet.Column(8).Width = 28.57 + 2.71;
            worksheet.Column(9).Width = 14.71 + 2.71;
            worksheet.Column(10).Width = 14.71 + 2.71;
            worksheet.Column(11).Width = 16.29 + 2.71;
            worksheet.Column(12).Width = 18.71 + 2.71;


            worksheet.Row(4).Height = 14.25;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:J3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ED7D31"));

        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
