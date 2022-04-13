using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato64_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato64_Model cotizacion)
        {
            byte[] file;

            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Bartolome");

                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(300, 200);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                
                
                string frase1 = "ADQUISICIÓN DE DISPOSITIVOS MEDICOS SOLICITADOS POR EL SERVICIO DE FARMACIA DEL HONADOMANI-SB";
               


                workSheet.Cells["A3"].Value = "Exp. 745-2021";
                workSheet.Cells["A3"].Style.Font.Size = 24;
                workSheet.Cells["A3"].Style.Font.Name = "Arial";
                workSheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A3"].Style.Font.Bold = true;
                workSheet.Cells["A3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FF0000"));

                workSheet.Cells["A5"].Value = "Cotiz.Nro. " + cotizacion.Prov_NroCotizacion + "-2022 Unilene S.A.C";
                workSheet.Cells["A5"].Style.Font.Size = 18;
                workSheet.Cells["A5"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A5"].Style.Font.Bold = true;
                workSheet.Cells["A5"].Style.Font.UnderLine = true;


                workSheet.Cells["L5"].Value = "FECHA";
                workSheet.Cells["L5"].Style.Font.Size = 28;
                workSheet.Cells["L5"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["L5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["L5"].Style.Font.Bold = true;

                workSheet.Cells["M5"].Value = cotizacion.Fecha_Cotizacion;
                workSheet.Cells["M5"].Style.Font.Size = 20;
                workSheet.Cells["M5"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["M5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M5"].Style.Numberformat.Format = "dd/MM/yyyy";

                workSheet.Cells["A7"].Value = "Señores:";
                workSheet.Cells["A7"].Style.Font.Size = 28;
                workSheet.Cells["A7"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A7"].Style.Font.Bold = true;


                workSheet.Cells["A8"].Value = "HOSPITAL NACIONAL DOCENTE MADRE NIÑO - SAN BARTOLOME";
                workSheet.Cells["A8"].Style.Font.Size = 28;
                workSheet.Cells["A8"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                workSheet.Cells["A8"].Style.Font.Bold = true;

                workSheet.Cells["A9"].Value = "Av. Alfonso Ugarte N° 825 -Cercado de Lima.";
                workSheet.Cells["A9"].Style.Font.Size = 20;
                workSheet.Cells["A9"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A9"].Style.Font.Bold = true;


                workSheet.Cells["A10"].Value = "Av. Alfonso Ugarte N° 825 -Cercado de Lima.";
                workSheet.Cells["A10"].Style.Font.Size = 28;
                workSheet.Cells["A10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A10"].Style.Font.Bold = true;


                workSheet.Cells["A12"].Value = "De  nuestra consideración:";
                workSheet.Cells["A12"].Style.Font.Size = 26;
                workSheet.Cells["A12"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["A13"].Value = "El  que  suscribe, Don Luis German Castillo Vasquez  identificado con  DNI  Nº 07227297 de la empresa: Unilene S.A.C., presento mi COTIZACIÓN que fue solicitada para la "+frase1 

                + " , la misma que presenta los siguientes detalles:";
                workSheet.Cells["A13"].Style.Font.Size = 26;
                workSheet.Cells["A13"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
               




                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);

                //DETALLE

                workSheet.Cells["A15"].Value = "ITEM";
                workSheet.Cells["A15"].Style.Font.Size = 25;
                workSheet.Cells["A15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A15"].Style.WrapText = true;

                workSheet.Cells["B15"].Value = "ITEM";
                workSheet.Cells["B15"].Style.Font.Size = 25;
                workSheet.Cells["B15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["B15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B15"].Style.WrapText = true;


                workSheet.Cells["C15"].Value = "DETALLE DEL BIEN";
                workSheet.Cells["C15"].Style.Font.Size = 25;
                workSheet.Cells["C15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["C15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C15"].Style.WrapText = true;

                workSheet.Cells["G15"].Value = "UNIDAD DE MEDIDA";
                workSheet.Cells["G15"].Style.Font.Size = 25;
                workSheet.Cells["G15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["G15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G15"].Style.WrapText = true;


                workSheet.Cells["H15"].Value = "CANT.";
                workSheet.Cells["H15"].Style.Font.Size = 25;
                workSheet.Cells["H15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["H15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H15"].Style.WrapText = true;

                workSheet.Cells["I15"].Value = "MARCA";
                workSheet.Cells["I15"].Style.Font.Size = 25;
                workSheet.Cells["I15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["I15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I15"].Style.WrapText = true;

                workSheet.Cells["J15"].Value = "MODELO";
                workSheet.Cells["J15"].Style.Font.Size = 25;
                workSheet.Cells["J15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["J15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J15"].Style.WrapText = true;

                workSheet.Cells["K15"].Value = "PROCEDENCIA";
                workSheet.Cells["K15"].Style.Font.Size = 25;
                workSheet.Cells["K15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["K15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K15"].Style.WrapText = true;

                workSheet.Cells["L15"].Value = "AÑO DE FABRICACIÓN";
                workSheet.Cells["L15"].Style.Font.Size = 25;
                workSheet.Cells["L15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["L15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L15"].Style.WrapText = true;

                workSheet.Cells["M15"].Value = "PRECIO UNITARIO S/ INC. IGV";
                workSheet.Cells["M15"].Style.Font.Size = 25;
                workSheet.Cells["M15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["M15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M15"].Style.WrapText = true;

                workSheet.Cells["N15"].Value = "PRECIO TOTAL S/ INC. IGV";
                workSheet.Cells["N15"].Style.Font.Size = 25;
                workSheet.Cells["N15"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["N15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N15"].Style.WrapText = true;

                int index = 16;

                foreach (Coti_Formato64_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Row(index).Height = 77.25;
                    workSheet.Cells["A" + index].Value = index - 15;
                    workSheet.Cells["A" + index].Style.Font.Size = 20;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.WrapText = true;
                    workSheet.Cells["A" + index].Style.Font.Bold = true;
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";


                    workSheet.Cells["B" + index].Value = item.Item;
                    workSheet.Cells["B" + index].Style.Font.Size = 12;
                    workSheet.Cells["B" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;
                    workSheet.Cells["B" + index].Style.Font.Bold = true;

                    workSheet.Cells["C" + index + ":F" + index].Value = item.Detalle;
                    workSheet.Cells["C" + index + ":F" + index].Style.Font.Size = 16;
                    workSheet.Cells["C" + index + ":F" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["C" + index + ":F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index + ":F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["C" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index + ":F" + index].Style.WrapText = true;
                    workSheet.Cells["C" + index + ":F" + index].Merge = true;


                    workSheet.Cells["G" + index].Value = item.UndMedida;
                    workSheet.Cells["G" + index].Style.Font.Size = 20;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;
                    workSheet.Cells["G" + index].Style.Font.Bold = true;



                    workSheet.Cells["H" + index].Value = item.Cantidad;
                    workSheet.Cells["H" + index].Style.Font.Size = 20;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;
                    workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells["H" + index].Style.Font.Bold = true;

                    workSheet.Cells["I" + index].Value = item.Marca;
                    workSheet.Cells["I" + index].Style.Font.Size = 20;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;


                    workSheet.Cells["J" + index].Value = item.Modelo;
                    workSheet.Cells["J" + index].Style.Font.Size = 20;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;


                    workSheet.Cells["K" + index].Value = item.Procedencia;
                    workSheet.Cells["K" + index].Style.Font.Size = 25;
                    workSheet.Cells["K" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;

                    workSheet.Cells["L" + index].Value = item.AnioFabricacion;
                    workSheet.Cells["L" + index].Style.Font.Size = 25;
                    workSheet.Cells["L" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;

                    workSheet.Cells["M" + index].Value = item.PreUnitario;
                    workSheet.Cells["M" + index].Style.Font.Size = 25;
                    workSheet.Cells["M" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;
                    workSheet.Cells["M" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["N" + index].Value = item.PreTotal;
                    workSheet.Cells["N" + index].Style.Font.Size = 25;
                    workSheet.Cells["N" + index].Style.Font.Name = "Arial Narrow";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;
                    workSheet.Cells["M" + index].Style.Numberformat.Format = "#,##0.00";
                    index++;
                }
                
                workSheet.Row(index).Height = 45.00;
                workSheet.Cells["A" + index + ":M" + index].Value = "T O T A L  G E N E R A L (INC IGV) ";
                workSheet.Cells["A" + index + ":M" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":M" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["A" + index + ":M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":M" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":M" + index].Merge = true;
                workSheet.Cells["A" + index + ":M" + index].Style.Font.Bold = true;

                workSheet.Cells["N" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["N" + index].Style.Font.Size = 25;
                workSheet.Cells["N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["N" + index].Style.WrapText = true;
                workSheet.Cells["N" + index].Merge = true;
                workSheet.Cells["N" + index].Style.Font.Bold = true;

                index++;
                workSheet.Row(index).Height = 17.25;
                index++;
                workSheet.Row(index).Height = 109.50;

                workSheet.Cells["A" + index + ":N" + index].Value = "Declaro que he revisado en forma detallada la documentación remitida y que nuestra cotización CUMPLE con las ESPECIFICACIONES TÉCNICAS, " +
                    "enviados e incluye todos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso, los costos laborales respectivos conforme a la legislación vigente, " +
                    "así como cualquier otro concepto que le sea aplicable y que pueda incidir sobre el valor del bien.";
                workSheet.Cells["A" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["A" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 16.5;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "RAZON SOCIAL";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;



                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "RUC";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "VALIDEZ DE LA COTIZACION";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_ValiCotizacion;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "PLAZO DE ENTREGA (DÍAS CALENDARIO)";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;


                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "FORMA DE PAGO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_FormaPago;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "PLAZO  DE GARANTIA";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_Garantia;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;


                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "VIGENCIA DEL PRODUCTO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_vigProducto;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;


                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "NOMBRE DE LA PERSONA DE CONTACTO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "N° DE TELÉFONO DE CONTACTO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_ContTelefono;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "CORREO ELECTRONICO DE CONTACTO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_ContEmail;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;


                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "N° CODIGO DE CUENA INTERBANCARIA (CCI)";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_Cbancaria;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Bold = true;

                index++;
                workSheet.Row(index).Height = 45.75;
                workSheet.Cells["A" + index + ":D" + index].Value = "INFORMACIÓN TÉCNICA ADICIONAL (SI/NO)";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Bold = true;

                workSheet.Cells["E" + index].Value = ":";
                workSheet.Cells["E" + index].Style.Font.Size = 25;
                workSheet.Cells["E" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":N" + index].Value = cotizacion.Prov_InfAdicional;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["F" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["F" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["F" + index + ":N" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 22.5;
                index++;
                workSheet.Row(index).Height = 68.25;
                index++;
                workSheet.Row(index).Height = 5.25;
                index++;
                workSheet.Row(index).Height = 29.25;
                index++;
                workSheet.Row(index).Height = 45.75;
                index++;
                workSheet.Row(index).Height = 32.25;
                workSheet.Cells["A" + index + ":N" + index].Value = "Represente/Representante Legal";
                workSheet.Cells["A" + index + ":N" + index].Style.Font.Size = 25;
                workSheet.Cells["A" + index + ":N" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":N" + index].Merge = true;
         


                workSheet.View.ZoomScale = 80;
                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }

            return reporte;
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Column(1).Width = 10.43 + 0.71;
            workSheet.Column(2).Width = 19.00 + 0.71;
            workSheet.Column(3).Width = 32.57 + 0.71;
            workSheet.Column(4).Width = 44.00 + 0.71;
            workSheet.Column(5).Width = 1.57 + 0.71;
            workSheet.Column(6).Width = 3.43 + 0.71;
            workSheet.Column(7).Width = 17.71 + 0.71;
            workSheet.Column(8).Width = 15.29 + 0.71;
            workSheet.Column(9).Width = 17.71 + 0.71;
            workSheet.Column(10).Width = 18.00 + 0.71;
            workSheet.Column(11).Width = 29.71 + 0.71;
            workSheet.Column(12).Width = 29.14 + 0.71;
            workSheet.Column(13).Width = 20.71 + 0.71;
            workSheet.Column(14).Width = 23.43 + 0.71;

            workSheet.Row(1).Height = 31.5;
            workSheet.Row(2).Height = 115.5;
            workSheet.Row(3).Height = 70.5;
            workSheet.Row(4).Height = 17.25;
            workSheet.Row(5).Height = 39.00;
            workSheet.Row(6).Height = 5.25;
            workSheet.Row(7).Height = 35.25;
            workSheet.Row(8).Height = 26.25;
            workSheet.Row(9).Height = 26.25;
            workSheet.Row(10).Height = 26.25;
            workSheet.Row(11).Height = 16.5;
            workSheet.Row(12).Height = 30.75;
            workSheet.Row(13).Height = 95.25;
            workSheet.Row(14).Height = 16.50;
            workSheet.Row(15).Height = 126.00;
            

        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A3:N3"].Merge = true;
            workSheet.Cells["A5:C5"].Merge = true;
            workSheet.Cells["M5:N5"].Merge = true;
            workSheet.Cells["A7:B7"].Merge = true;
            workSheet.Cells["A8:H8"].Merge = true;
            workSheet.Cells["A9:H9"].Merge = true;
            workSheet.Cells["A10:D10"].Merge = true;
            workSheet.Cells["A12:N12"].Merge = true;
            workSheet.Cells["A13:N13"].Merge = true;
            workSheet.Cells["C15:F15"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["M5:N5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C15:F15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void TextoNegrita(ExcelWorksheet workSheet)
        {

        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A15:N15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
        }
    }
}
