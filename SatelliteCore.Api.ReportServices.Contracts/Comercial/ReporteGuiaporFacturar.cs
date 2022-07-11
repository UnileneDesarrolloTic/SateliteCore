using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Comercial
{
    public class ReporteGuiaporFacturar
    {
        public string ExportarListarGuiaPorFactura(IEnumerable<FormatoGuiaPorFacturarModel> ListaGuiaPorFactura, DatosEstructuraGuiaPorFacturarModel dato)
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
                UnirCeldas(worksheet);
                PintarCeldas(worksheet);
                BordesCeldas(worksheet);
                BordesCeldas(worksheet);
                TextoNegrita(worksheet);

                string titulo = "";
                if (dato.Tipo == "GF")
                {
                    titulo = "REPORTE GENERAL DE GUIAS PENDIENTES DE FACTURAR";
                }
                else
                {
                    titulo = "REPORTE GENERAL DE GUIAS FACTURADAS";
                }


                worksheet.Cells["A1"].Value = titulo;
                worksheet.Cells["A1"].Style.Font.Size = 14;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                worksheet.Cells["A3"].Value = "SERIE";
                worksheet.Cells["A3"].Style.Font.Size = 10;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["B3"].Value = "NRO_GUIA";
                worksheet.Cells["B3"].Style.Font.Size = 10;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["C3"].Value = "FECHA_GUIA";
                worksheet.Cells["C3"].Style.Font.Size = 10;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D3"].Value = "E.COMERCIAL";
                worksheet.Cells["D3"].Style.Font.Size = 10;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["E3"].Value = "U.COMERCIAL";
                worksheet.Cells["E3"].Style.Font.Size = 10;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["F3"].Value = "FACTURA";
                worksheet.Cells["F3"].Style.Font.Size = 10;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["G3"].Value = "F.FACTURA";
                worksheet.Cells["G3"].Style.Font.Size = 10;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["H3"].Value = "PERSONA";
                worksheet.Cells["H3"].Style.Font.Size = 10;
                worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["I3"].Value = "RUC";
                worksheet.Cells["I3"].Style.Font.Size = 10;
                worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["J3"].Value = "NOMBRE_CLIENTE";
                worksheet.Cells["J3"].Style.Font.Size = 10;
                worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["K3"].Value = "ESTADO";
                worksheet.Cells["K3"].Style.Font.Size = 10;
                worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["L3"].Value = "ULTIMO USUARIO";
                worksheet.Cells["L3"].Style.Font.Size = 10;
                worksheet.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["M3"].Value = "PROCESO";
                worksheet.Cells["M3"].Style.Font.Size = 10;
                worksheet.Cells["M3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["N3"].Value = "E/PECOSA";
                worksheet.Cells["N3"].Style.Font.Size = 10;
                worksheet.Cells["N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["O3"].Value = "E.LOGISTICA";
                worksheet.Cells["O3"].Style.Font.Size = 10;
                worksheet.Cells["O3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int row = 4;


                foreach (FormatoGuiaPorFacturarModel item in ListaGuiaPorFactura)
                {
                    worksheet.Row(row).Height = 25.5;




                    worksheet.Cells["A" + row].Value = item.SerieNumero;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["B" + row].Value = item.GuiaNumero;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                    worksheet.Cells["C" + row].Value = item.FechaDocumento;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["D" + row].Value = item.ComentariosEntrega == true ? "SI" : "NO";
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["E" + row].Value = item.UsuComercial;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["F" + row].Value = item.FacturaNumero;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                    worksheet.Cells["G" + row].Value = item.FacturaFecha;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["H" + row].Value = item.Destinatario;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["I" + row].Value = item.DestinatarioRUC;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["J" + row].Value = item.DestinatarioNombre;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + row].Style.Font.Size = 10;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                    worksheet.Cells["K" + row].Value = item.EstadoGuia;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["K" + row].Style.Font.Size = 10;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["L" + row].Value = item.UltimoUsuario;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["L" + row].Style.Font.Size = 10;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["M" + row].Value = item.LicitacionNumeroProceso;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["M" + row].Style.Font.Size = 10;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["N" + row].Value = item.ReprogramacionPuntoPartida;
                    worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["N" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["N" + row].Style.Font.Size = 10;
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["O" + row].Value = item.EstadoLogistica;
                    worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["O" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["O" + row].Style.Font.Size = 10;
                    worksheet.Cells["O" + row].Style.WrapText = true;
                    worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    if (item.ComentariosEntrega)
                    {
                        worksheet.Cells["A" + row + ":O" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#60D460"));
                    }

                    row++;
                }


                if (dato.Tipo == "GF")
                {
                    worksheet.Column(6).Hidden = true;
                    worksheet.Column(7).Hidden = true;
                }
                else
                {
                    worksheet.Column(6).Hidden = false;
                    worksheet.Column(7).Hidden = false;
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
            worksheet.Column(1).Width = 4.43 + 2.71;
            worksheet.Column(2).Width = 8.43 + 2.71;
            worksheet.Column(3).Width = 10.43 + 2.71;
            worksheet.Column(4).Width = 10.43 + 2.71;
            worksheet.Column(5).Width = 10.43 + 2.71;
            worksheet.Column(6).Width = 10.43 + 2.71;
            worksheet.Column(7).Width = 9.86 + 2.71;
            worksheet.Column(8).Width = 7.43 + 2.71;
            worksheet.Column(9).Width = 11.29 + 2.71;
            worksheet.Column(10).Width = 45 + 2.71;
            worksheet.Column(11).Width = 6.29 + 2.71;
            worksheet.Column(12).Width = 8.43 + 2.71;
            worksheet.Column(13).Width = 8.43 + 2.71;
            worksheet.Column(14).Width = 8.43 + 2.71;
            worksheet.Column(15).Width = 10.43 + 2.71;

        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1:N1"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {
            // worksheet.Cells["N14,J5"].Style.Font.UnderLine = true;
            worksheet.Cells["A3:O3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D8D8D8"));

        }
        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
