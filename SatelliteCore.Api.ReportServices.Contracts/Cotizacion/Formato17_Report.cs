using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato17_Report
    {

        public static string Exportar(Image logoUnilene, Coti_Formato17_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Sabogal");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(100, 50);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);



                workSheet.Cells["I1"].Value = "Cotización "+ cotizacion.Prov_NroCotizacion + " - Unilene SAC";
                workSheet.Cells["I1"].Style.Font.Size = 9;
                workSheet.Cells["I1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["I1"].Style.Font.Bold = true;
                workSheet.Cells["I1"].Style.WrapText = true;

                workSheet.Cells["A3"].Value = "FORMATO DE COTIZACION DE BIENES";
                workSheet.Cells["A3"].Style.Font.Size = 10;
                workSheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A3"].Style.Font.Bold = true;

                workSheet.Cells["A4"].Value = "Lima " + cotizacion.Fecha_Cotizacion.ToLongDateString();
                workSheet.Cells["A4"].Style.Font.Size = 10;
                workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["A7"].Value = "Señores: ";
                workSheet.Cells["A7"].Style.Font.Size = 8;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A7"].Style.Font.Bold = true;

                workSheet.Cells["B7"].Value = "SEGURO SOCIAL DE SALUD ESSALUD - RED ASISTENCIAL ICA";
                workSheet.Cells["B7"].Style.Font.Size = 8;
                workSheet.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B7"].Style.Font.Bold = true;

                workSheet.Cells["A9"].Value = "De mi consideración:";
                workSheet.Cells["A9"].Style.Font.Size = 10;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["A10"].Value = "En respuesta a la solicitud de cotización sobre la “Adquisición  de Delegados”, y después de haber analizado las especificaciones técnicas del mencionado requerimiento, las mismas que acepto en todos sus extremos, indico que cumplo con TODOS los requerimientos solicitados. \n"+
                                               "Asimismo, declaro que las características técnicas de los bienes cotizados por mi representada se ajustan a lo requerido por entidad. En tal sentido, indico que el costo total por la solución requerida es la que detallo a continuación: solución requerida es la que detallo a continuación:";
                workSheet.Cells["A10"].Style.Font.Size = 10;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A10"].Style.WrapText = true;

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);
                PintarCeldas(workSheet);

                //DETALLE

                workSheet.Cells["A15"].Value = "Ítem";
                workSheet.Cells["A15"].Style.Font.Size = 8;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A15"].Style.Font.Bold = true;
                workSheet.Cells["A15"].Style.WrapText = true;

                workSheet.Cells["B15"].Value = "Código del Material";
                workSheet.Cells["B15"].Style.Font.Size = 8;
                workSheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B15"].Style.Font.Bold = true;
                workSheet.Cells["B15"].Style.WrapText = true;

                workSheet.Cells["C15"].Value = "DESCRIPCION";
                workSheet.Cells["C15"].Style.Font.Size = 8;
                workSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C15"].Style.Font.Bold = true;
                workSheet.Cells["C15"].Style.WrapText = true;

                workSheet.Cells["D15"].Value = "CANT. UNITARIA";
                workSheet.Cells["D15"].Style.Font.Size = 8;
                workSheet.Cells["D15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D15"].Style.Font.Bold = true;
                workSheet.Cells["D15"].Style.WrapText = true;

                workSheet.Cells["E15"].Value = "Unidad de Medida";
                workSheet.Cells["E15"].Style.Font.Size = 8;
                workSheet.Cells["E15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E15"].Style.Font.Bold = true;
                workSheet.Cells["E15"].Style.WrapText = true;

                workSheet.Cells["F15"].Value = "Marca";
                workSheet.Cells["F15"].Style.Font.Size = 8;
                workSheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F15"].Style.Font.Bold = true;
                workSheet.Cells["F15"].Style.WrapText = true;


                workSheet.Cells["G15"].Value = "Modelo";
                workSheet.Cells["G15"].Style.Font.Size = 8;
                workSheet.Cells["G15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G15"].Style.Font.Bold = true;
                workSheet.Cells["G15"].Style.WrapText = true;

                workSheet.Cells["H15"].Value = "País de Procedencia";
                workSheet.Cells["H15"].Style.Font.Size = 8;
                workSheet.Cells["H15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H15"].Style.Font.Bold = true;
                workSheet.Cells["H15"].Style.WrapText = true;

                workSheet.Cells["I15"].Value = "PRECIO UNITARIO";
                workSheet.Cells["I15"].Style.Font.Size = 8;
                workSheet.Cells["I15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I15"].Style.Font.Bold = true;
                workSheet.Cells["I15"].Style.WrapText = true;

                workSheet.Cells["J15"].Value = "TOTAL";
                workSheet.Cells["J15"].Style.Font.Size = 8;
                workSheet.Cells["J15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J15"].Style.Font.Bold = true;
                workSheet.Cells["J15"].Style.WrapText = true;



                int index = 16;
                foreach (Coti_Formato17_Detalle item in cotizacion.Detalle)
                {

                    workSheet.Row(index).Height = 39.75;
                    workSheet.Cells["A" + index].Value = index - 15;
                    workSheet.Cells["A" + index].Style.Font.Size = 10;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.WrapText = true;
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                    workSheet.Cells["B" + index].Value = item.CodMaterial;
                    workSheet.Cells["B" + index].Style.Font.Size = 10;
                    workSheet.Cells["B" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;

                    workSheet.Cells["C" + index].Value = item.Descripcion;
                    workSheet.Cells["C" + index].Style.Font.Size = 10;
                    workSheet.Cells["C" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;


                    workSheet.Cells["D" + index].Value = item.Cantidad;
                    workSheet.Cells["D" + index].Style.Font.Size = 9;
                    workSheet.Cells["D" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;
                    workSheet.Cells["D" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["E" + index].Value = item.UndMedida;
                    workSheet.Cells["E" + index].Style.Font.Size = 9;
                    workSheet.Cells["E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;

                    workSheet.Cells["F" + index].Value = item.Marca;
                    workSheet.Cells["F" + index].Style.Font.Size = 9;
                    workSheet.Cells["F" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;


                    workSheet.Cells["G" + index].Value = item.Modelo;
                    workSheet.Cells["G" + index].Style.Font.Size = 9;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;

                    workSheet.Cells["H" + index].Value = item.Procedencia;
                    workSheet.Cells["H" + index].Style.Font.Size = 9;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;

                    workSheet.Cells["I" + index].Value = item.PreUnitario;
                    workSheet.Cells["I" + index].Style.Font.Size = 9;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;
                    workSheet.Cells["I" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["J" + index].Value = item.PreTotal;
                    workSheet.Cells["J" + index].Style.Font.Size = 9;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;
                    workSheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";

                    index++;

                }



                workSheet.Row(index).Height = 17;
                workSheet.Cells["A" + index + ":H" + index].Merge = true;
                workSheet.Cells["A" + index + ":H" + index].Value = "VALOR TOTAL DE LA COTIZACIÓN";
                workSheet.Cells["A" + index + ":H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                workSheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";

                workSheet.Cells["J" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["J" + index].Style.Font.Bold = true;
                workSheet.Cells["J" + index].Style.Font.Size = 10;
                workSheet.Cells["J" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 8;

                index++;
                workSheet.Cells["A" + index + ":J" + (index + 5)].Merge = true;
                workSheet.Cells["A" + index + ":J" + (index + 5)].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, inspecciones, pruebas y , de ser el caso, los costos laborales conf orme la legislación v igente, así como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y /o serv icio a contratar; excepto la de aquellos prov eedores que gocen de alguna exoneración legal, no incluirán en el precio de su of erta los tributos respectiv os. Asimismo, declaro bajo juramente que, mi persona y /o mi representada no cuenta con impedimentos para contratar con el Estado, conf orme lo establece el artículo 11 del Texto Único Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado, aprobado por Decreto Supremo N° 082-2019-EF";
                workSheet.Cells["A" + index + ":J" + (index + 5)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["A" + index + ":J" + (index + 5)].Style.Font.Size = 9;
                workSheet.Cells["A" + index + ":J" + (index + 5)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":J" + (index + 5)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Merge = true;
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Value = "Razón Social";
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + (index + 6) + ":C" + (index + 6)].Style.WrapText = true;


                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Merge = true;
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 6) + ":J" + (index + 6)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Merge = true;
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Value = "N° RUC";
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + (index + 7) + ":C" + (index + 7)].Style.WrapText = true;


                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Merge = true;
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 7) + ":J" + (index + 7)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Merge = true;
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Value = "Plazo de entrega";
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + (index + 8) + ":C" + (index + 8)].Style.WrapText = true;


                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Merge = true;
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 8) + ":J" + (index + 8)].Style.WrapText = true;



                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Merge = true;
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Value = "Forma de pago";
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + (index + 9) + ":C" + (index + 9)].Style.WrapText = true;


                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Merge = true;
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 9) + ":J" + (index + 9)].Style.WrapText = true;

                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Merge = true;
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Value = "Garantía";
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + (index + 10) + ":C" + (index + 10)].Style.WrapText = true;


                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Merge = true;
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Value = cotizacion.Prov_Garantia;
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 10) + ":J" + (index + 10)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Merge = true;
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Value = "Correo electrónico";
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Style.Font.Name = "Arial";
                workSheet.Cells["A" + (index + 11) + ":C" + (index + 11)].Style.WrapText = true;


                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Merge = true;
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Value = cotizacion.Prov_Email;
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 11) + ":J" + (index + 11)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Merge = true;
                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Value = "Teléfono Fijo";
                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 12) + ":C" + (index + 12)].Style.Font.Name = "Arial";

                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Merge = true;
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Value = cotizacion.Prov_Email;
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 12) + ":J" + (index + 12)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Merge = true;
                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Value = "Persona de contacto";
                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 13) + ":C" + (index + 13)].Style.Font.Name = "Arial";

                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Merge = true;
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 13) + ":J" + (index + 13)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Merge = true;
                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Value = "Teléfono Móvil";
                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 14) + ":C" + (index + 14)].Style.Font.Name = "Arial";

                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Merge = true;
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Value = cotizacion.Prov_ContTelefono;
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 14) + ":J" + (index + 14)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Merge = true;
                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Value = "Vigencia de la oferta";
                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 15) + ":C" + (index + 15)].Style.Font.Name = "Arial";

                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Merge = true;
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 15) + ":J" + (index + 15)].Style.WrapText = true;


                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Merge = true;
                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Style.Font.Bold = true;
                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Value = "Registro Sanitario";
                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Style.Font.Size = 10;
                workSheet.Cells["A" + (index + 16) + ":C" + (index + 16)].Style.Font.Name = "Arial";

                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Merge = true;
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Style.Font.Bold = true;
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Value = cotizacion.Prov_RegSanitario;
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Style.Font.Size = 10;
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Style.Font.Name = "Arial";
                workSheet.Cells["D" + (index + 16) + ":J" + (index + 16)].Style.WrapText = true;



                workSheet.Row(index + 22).Height = 27.5;
                workSheet.Cells["C" + (index + 22) + ":F" + (index + 22)].Merge = true;
                workSheet.Cells["C" + (index + 22) + ":F" + (index + 22)].Value = "Firma, nombres y apellidos del proveedor o representante legal o Persona autorizada para emitir cotizaciones";
                workSheet.Cells["C" + (index + 22) + ":F" + (index + 22)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + (index + 22) + ":F" + (index + 22)].Style.Font.Size = 10;
                workSheet.Cells["C" + (index + 22) + ":F" + (index + 22)].Style.Font.Name = "Arial";
                workSheet.Cells["C" + (index + 22) + ":F" + (index + 22)].Style.WrapText = true;



                workSheet.View.ZoomScale = 85;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }

            return reporte;
        }


        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Column(1).Width = 7.29 + 0.71;
            workSheet.Column(2).Width = 9.71 + 0.71;
            workSheet.Column(3).Width = 26.14 + 0.71;
            workSheet.Column(4).Width = 9.57 + 0.71;
            workSheet.Column(5).Width = 8.86 + 0.71;

            workSheet.Column(6).Width = 8.57 + 0.71;
            workSheet.Column(7).Width = 10.71 + 0.71;
            workSheet.Column(8).Width = 10.57 + 0.71;
            workSheet.Column(9).Width = 8.57 + 0.71;
            workSheet.Column(10).Width = 11.57 + 0.71;



            workSheet.Row(1).Height = 26.5;
            workSheet.Row(2).Height = 6.5;
            workSheet.Row(3).Height = 17;
            workSheet.Row(4).Height = 15.5;
            workSheet.Row(5).Height = 5.00;

            workSheet.Row(6).Height = 12.75;
            workSheet.Row(7).Height = 13.25;
            workSheet.Row(8).Height = 13.25;

            workSheet.Row(9).Height = 13.25;
            workSheet.Row(10).Height = 12.75;
            workSheet.Row(11).Height = 00.00;
            workSheet.Row(12).Height = 12.75;

            workSheet.Row(13).Height = 38.25;
            workSheet.Row(14).Height = 12.50;
            workSheet.Row(15).Height = 29.00;

        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["I1:J1"].Merge = true;
            workSheet.Cells["A3:J3"].Merge = true;
            workSheet.Cells["A4:E4"].Merge = true;
            workSheet.Cells["B7:E7"].Merge = true;

            workSheet.Cells["A9:C9"].Merge = true;
            workSheet.Cells["A10:J13"].Merge = true;

        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
     

            workSheet.Cells["A15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A15:J15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BC2E6"));
        }
    }
}
