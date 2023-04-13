using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;


namespace SatelliteCore.Api.ReportServices.Contracts.Produccion
{
    public class ReporteCompraAguja_Excel
    {
        public string GenerarReporte(IEnumerable<DatosFormatoListadoSeguimientoCompraAguja> dato)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Compra Drogueria");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);

                TextoNegrita(worksheet);

           

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
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void pintarCabecera(ExcelWorksheet worksheet, int fila)
        {
        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }


    }
}
