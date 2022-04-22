using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato27_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato27_Model cotizacion)
        {

            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Piura");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(130, 50);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                PintarCeldas(worksheet);
                BordesCeldas(worksheet);
                BordesCeldas(worksheet);
                TextoNegrita(worksheet);


                worksheet.Cells["I1"].Value = "Cotiz Nº " + cotizacion.Prov_NroCotizacion + "-2022-Unilene S.A.C";
                worksheet.Cells["I1"].Style.Font.Size = 10;
                worksheet.Cells["I1"].Style.WrapText = true;
                worksheet.Cells["I1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A3"].Value = "FORMATO DE COTIZACION DE BIENES";
                worksheet.Cells["A3"].Style.Font.Size = 10;
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A4"].Value = "Lima " + cotizacion.Prov_Fecha.ToLongDateString();
                worksheet.Cells["A4"].Style.Font.Size = 10;
                worksheet.Cells["A4"].Style.WrapText = true;
                worksheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A7"].Value = "SEGURO SOCIAL DE SALUD ESSALUD - ESSALUD PIURA";
                worksheet.Cells["A7"].Style.Font.Size = 10;
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A7"].Style.WrapText = true;
                worksheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A9"].Value = "De mi consideración:";
                worksheet.Cells["A9"].Style.Font.Size = 10;
                worksheet.Cells["A9"].Style.WrapText = true;
                worksheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A10"].Value = "En respuesta a la solicitud de cotización sobre Disp. Médicos Delegados a compra local. ESSALUD Piura y después de haber analizado las especificaciones técnicas del mencionado requerimiento, las mismas que acepto en todos sus extremos, indico que cumplo " +
                    "con TODOS los requerimientos solicitados.\nAsimismo, declaro las características técnicas de los bienes cotizados por mi representada se ajustan a lo requerido por su Entidad.En tal sentido, indico que el costo total por la solución requerida es la que detallo a continuación: ";
                worksheet.Cells["A10"].Style.Font.Size = 10;
                worksheet.Cells["A10"].Style.WrapText = true;
                worksheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;



                worksheet.Cells["A15"].Value = "Ítem";
                worksheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A15"].Style.Font.Bold = true;
                worksheet.Cells["A15"].Style.Font.Size = 8;
                worksheet.Cells["A15"].Style.WrapText = true;
                worksheet.Cells["A15"].Style.Font.Name = "Arial";

                worksheet.Cells["B15"].Value = "Código del Material";
                worksheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B15"].Style.Font.Bold = true;
                worksheet.Cells["B15"].Style.Font.Size = 8;
                worksheet.Cells["B15"].Style.WrapText = true;
                worksheet.Cells["B15"].Style.Font.Name = "Arial";

                worksheet.Cells["C15"].Value = "DESCRIPCION";
                worksheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C15"].Style.Font.Bold = true;
                worksheet.Cells["C15"].Style.Font.Size = 8;
                worksheet.Cells["C15"].Style.WrapText = true;
                worksheet.Cells["C15"].Style.Font.Name = "Arial";

                worksheet.Cells["D15"].Value = "CANT. UNITARIA";
                worksheet.Cells["D15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D15"].Style.Font.Bold = true;
                worksheet.Cells["D15"].Style.Font.Size = 8;
                worksheet.Cells["D15"].Style.WrapText = true;
                worksheet.Cells["D15"].Style.Font.Name = "Arial";

                worksheet.Cells["E15"].Value = "Unidad de Medida";
                worksheet.Cells["E15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E15"].Style.Font.Bold = true;
                worksheet.Cells["E15"].Style.Font.Size = 8;
                worksheet.Cells["E15"].Style.WrapText = true;
                worksheet.Cells["E15"].Style.Font.Name = "Arial";


                worksheet.Cells["F15"].Value = "Marca";
                worksheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F15"].Style.Font.Bold = true;
                worksheet.Cells["F15"].Style.Font.Size = 8;
                worksheet.Cells["F15"].Style.WrapText = true;
                worksheet.Cells["F15"].Style.Font.Name = "Arial";


                worksheet.Cells["G15"].Value = "Modelo";
                worksheet.Cells["G15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G15"].Style.Font.Bold = true;
                worksheet.Cells["G15"].Style.Font.Size = 8;
                worksheet.Cells["G15"].Style.WrapText = true;
                worksheet.Cells["G15"].Style.Font.Name = "Arial";

                worksheet.Cells["H15"].Value = "País de Procedencia";
                worksheet.Cells["H15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H15"].Style.Font.Bold = true;
                worksheet.Cells["H15"].Style.Font.Size = 8;
                worksheet.Cells["H15"].Style.WrapText = true;
                worksheet.Cells["H15"].Style.Font.Name = "Arial";

                worksheet.Cells["I15"].Value = "PRECIO UNITARIO";
                worksheet.Cells["I15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I15"].Style.Font.Bold = true;
                worksheet.Cells["I15"].Style.Font.Size = 8;
                worksheet.Cells["I15"].Style.WrapText = true;
                worksheet.Cells["I15"].Style.Font.Name = "Arial";

                worksheet.Cells["J15"].Value = "TOTAL";
                worksheet.Cells["J15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J15"].Style.Font.Bold = true;
                worksheet.Cells["J15"].Style.Font.Size = 8;
                worksheet.Cells["J15"].Style.WrapText = true;
                worksheet.Cells["J15"].Style.Font.Name = "Arial";



                int index = 16;
                foreach (Coti_Formato27_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Row(index).Height = 37.25;
                    worksheet.Cells["A" + index].Value = index - 15;
                    worksheet.Cells["A" + index].Style.Font.Size = 9;
                    worksheet.Cells["A" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + index].Style.WrapText = true;
                    worksheet.Cells["A" + index].Style.Numberformat.Format = "0";

                    worksheet.Cells["B" + index].Value = item.CodiMaterial;
                    worksheet.Cells["B" + index].Style.Font.Size = 9;
                    worksheet.Cells["B" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + index].Style.WrapText = true;

                    worksheet.Cells["C" + index].Value = item.Descripcion;
                    worksheet.Cells["C" + index].Style.Font.Size = 9;
                    worksheet.Cells["C" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + index].Style.WrapText = true;

                    worksheet.Cells["D" + index].Value = item.Cantidad;
                    worksheet.Cells["D" + index].Style.Font.Size = 9;
                    worksheet.Cells["D" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + index].Style.WrapText = true;
                    worksheet.Cells["D" + index].Style.Numberformat.Format = "#,##0";

                    worksheet.Cells["E" + index].Value = item.UnidMedida;
                    worksheet.Cells["E" + index].Style.Font.Size = 9;
                    worksheet.Cells["E" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + index].Style.WrapText = true;

                    worksheet.Cells["F" + index].Value = item.Marca;
                    worksheet.Cells["F" + index].Style.Font.Size = 9;
                    worksheet.Cells["F" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + index].Style.WrapText = true;

                    worksheet.Cells["G" + index].Value = item.Modelo;
                    worksheet.Cells["G" + index].Style.Font.Size = 9;
                    worksheet.Cells["G" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + index].Style.WrapText = true;

                    worksheet.Cells["H" + index].Value = item.Paisprocedencia;
                    worksheet.Cells["H" + index].Style.Font.Size = 9;
                    worksheet.Cells["H" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + index].Style.WrapText = true;

                    worksheet.Cells["I" + index].Value = item.PreUnitario;
                    worksheet.Cells["I" + index].Style.Font.Size = 9;
                    worksheet.Cells["I" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + index].Style.WrapText = true;
                    worksheet.Cells["I" + index].Style.Numberformat.Format = "#,##0.00";


                    worksheet.Cells["J" + index].Value = item.PreTotal;
                    worksheet.Cells["J" + index].Style.Font.Size = 9;
                    worksheet.Cells["J" + index].Style.Font.Name = "Arial";
                    worksheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + index].Style.WrapText = true;
                    worksheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";



                    index++;
                }

                worksheet.Cells["B" + index + ":" + "H" + index].Merge = true;
                worksheet.Cells["B" + index + ":" + "H" + index].Value = "TOTAL S/";
                worksheet.Cells["B" + index + ":" + "H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + index + ":" + "H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["J" + index].Value = cotizacion.Prov_ValorTotal;
                worksheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";

                index++;
                worksheet.Row(index).Height = 8.00;


                worksheet.Cells["A" + index + ":" + "J" + (index + 5)].Merge = true;
                worksheet.Cells["A" + index + ":" + "J" + (index + 5)].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento" +
                    " e incluy e todos los tributos, seguros, transporte, inspecciones, pruebas y , de ser el caso, los costos laborales conf orme la legislación v igente, así como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y /o serv icio a contratar;" +
                    " excepto la de aquellos prov eedores que gocen de alguna exoneración legal, no incluirán en el precio de su of erta los tributos respectiv os. Asimismo, declaro bajo juramente que, mi persona y /o mi representada no cuenta con impedimentos para contratar con el Estado, conf orme lo establece el artículo 11 del Texto Único Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado, " +
                    "aprobado por Decreto Supremo N° 082-2019-EF";
                worksheet.Cells["A" + index + ":" + "J" + (index + 5)].Style.WrapText = true;
                worksheet.Cells["A" + index + ":" + "J" + (index + 5)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + index + ":" + "J" + (index + 5)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                worksheet.Cells["A" + index + ":" + "J" + (index + 5)].Style.Font.Size = 9;

                worksheet.Row(index + 6).Height = 18.5;



                worksheet.Cells["A" + (index + 7) + ":" + "C" + (index + 7)].Merge = true;
                worksheet.Cells["A" + (index + 7) + ":" + "C" + (index + 7)].Value = "Razón Social";
                worksheet.Cells["A" + (index + 7) + ":" + "C" + (index + 7)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 7) + ":" + "C" + (index + 7)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 7) + ":" + "C" + (index + 7)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 7) + ":" + "C" + (index + 7)].Style.Font.Size = 9;

                worksheet.Cells["D" + (index + 7) + ":" + "J" + (index + 7)].Merge = true;
                worksheet.Cells["D" + (index + 7) + ":" + "J" + (index + 7)].Value = cotizacion.Prov_RazonSocial;
                worksheet.Cells["D" + (index + 7) + ":" + "J" + (index + 7)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 7) + ":" + "J" + (index + 7)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 7) + ":" + "J" + (index + 7)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 7) + ":" + "J" + (index + 7)].Style.Font.Size = 9;



                worksheet.Cells["A" + (index + 8) + ":" + "C" + (index + 8)].Merge = true;
                worksheet.Cells["A" + (index + 8) + ":" + "C" + (index + 8)].Value = "N° RUC";
                worksheet.Cells["A" + (index + 8) + ":" + "C" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 8) + ":" + "C" + (index + 8)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 8) + ":" + "C" + (index + 8)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 8) + ":" + "C" + (index + 8)].Style.Font.Size = 9;

                worksheet.Cells["D" + (index + 8) + ":" + "J" + (index + 8)].Merge = true;
                worksheet.Cells["D" + (index + 8) + ":" + "J" + (index + 8)].Value = cotizacion.Prov_Ruc;
                worksheet.Cells["D" + (index + 8) + ":" + "J" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 8) + ":" + "J" + (index + 8)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 8) + ":" + "J" + (index + 8)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 8) + ":" + "J" + (index + 8)].Style.Font.Size = 9;

                worksheet.Cells["A" + (index + 9) + ":" + "C" + (index + 9)].Merge = true;
                worksheet.Cells["A" + (index + 9) + ":" + "C" + (index + 9)].Value = "Plazo de entrega";
                worksheet.Cells["A" + (index + 9) + ":" + "C" + (index + 9)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 9) + ":" + "C" + (index + 9)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 9) + ":" + "C" + (index + 9)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 9) + ":" + "C" + (index + 9)].Style.Font.Size = 9;

                worksheet.Cells["D" + (index + 9) + ":" + "J" + (index + 9)].Merge = true;
                worksheet.Cells["D" + (index + 9) + ":" + "J" + (index + 9)].Value = cotizacion.Prov_PlazoEntrega;
                worksheet.Cells["D" + (index + 9) + ":" + "J" + (index + 9)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 9) + ":" + "J" + (index + 9)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 9) + ":" + "J" + (index + 9)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 9) + ":" + "J" + (index + 9)].Style.Font.Size = 9;

                worksheet.Cells["A" + (index + 10) + ":" + "C" + (index + 10)].Merge = true;
                worksheet.Cells["A" + (index + 10) + ":" + "C" + (index + 10)].Value = "Forma de pago";
                worksheet.Cells["A" + (index + 10) + ":" + "C" + (index + 10)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 10) + ":" + "C" + (index + 10)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 10) + ":" + "C" + (index + 10)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 10) + ":" + "C" + (index + 10)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 10) + ":" + "J" + (index + 10)].Merge = true;
                worksheet.Cells["D" + (index + 10) + ":" + "J" + (index + 10)].Value = cotizacion.Prov_FormaPago;
                worksheet.Cells["D" + (index + 10) + ":" + "J" + (index + 10)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 10) + ":" + "J" + (index + 10)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 10) + ":" + "J" + (index + 10)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 10) + ":" + "J" + (index + 10)].Style.Font.Size = 9;



                worksheet.Cells["A" + (index + 11) + ":" + "C" + (index + 11)].Merge = true;
                worksheet.Cells["A" + (index + 11) + ":" + "C" + (index + 11)].Value = "Garantía";
                worksheet.Cells["A" + (index + 11) + ":" + "C" + (index + 11)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 11) + ":" + "C" + (index + 11)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 11) + ":" + "C" + (index + 11)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 11) + ":" + "C" + (index + 11)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 11) + ":" + "J" + (index + 11)].Merge = true;
                worksheet.Cells["D" + (index + 11) + ":" + "J" + (index + 11)].Value = cotizacion.Prov_Garantia;
                worksheet.Cells["D" + (index + 11) + ":" + "J" + (index + 11)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 11) + ":" + "J" + (index + 11)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 11) + ":" + "J" + (index + 11)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 11) + ":" + "J" + (index + 11)].Style.Font.Size = 9;


                worksheet.Cells["A" + (index + 12) + ":" + "C" + (index + 12)].Merge = true;
                worksheet.Cells["A" + (index + 12) + ":" + "C" + (index + 12)].Value = "Correo electrónico";
                worksheet.Cells["A" + (index + 12) + ":" + "C" + (index + 12)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 12) + ":" + "C" + (index + 12)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 12) + ":" + "C" + (index + 12)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 12) + ":" + "C" + (index + 12)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 12) + ":" + "J" + (index + 12)].Merge = true;
                worksheet.Cells["D" + (index + 12) + ":" + "J" + (index + 12)].Value = cotizacion.Prov_Garantia;
                worksheet.Cells["D" + (index + 12) + ":" + "J" + (index + 12)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 12) + ":" + "J" + (index + 12)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 12) + ":" + "J" + (index + 12)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 12) + ":" + "J" + (index + 12)].Style.Font.Size = 9;


                worksheet.Cells["A" + (index + 13) + ":" + "C" + (index + 13)].Merge = true;
                worksheet.Cells["A" + (index + 13) + ":" + "C" + (index + 13)].Value = "Teléfono Fijo";
                worksheet.Cells["A" + (index + 13) + ":" + "C" + (index + 13)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 13) + ":" + "C" + (index + 13)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 13) + ":" + "C" + (index + 13)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 13) + ":" + "C" + (index + 13)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 13) + ":" + "J" + (index + 13)].Merge = true;
                worksheet.Cells["D" + (index + 13) + ":" + "J" + (index + 13)].Value = cotizacion.Prov_Telefono;
                worksheet.Cells["D" + (index + 13) + ":" + "J" + (index + 13)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 13) + ":" + "J" + (index + 13)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 13) + ":" + "J" + (index + 13)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 13) + ":" + "J" + (index + 13)].Style.Font.Size = 9;


                worksheet.Cells["A" + (index + 14) + ":" + "C" + (index + 14)].Merge = true;
                worksheet.Cells["A" + (index + 14) + ":" + "C" + (index + 14)].Value = "Persona de contacto";
                worksheet.Cells["A" + (index + 14) + ":" + "C" + (index + 14)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 14) + ":" + "C" + (index + 14)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 14) + ":" + "C" + (index + 14)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 14) + ":" + "C" + (index + 14)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 14) + ":" + "J" + (index + 14)].Merge = true;
                worksheet.Cells["D" + (index + 14) + ":" + "J" + (index + 14)].Value = cotizacion.Prov_Contacto;
                worksheet.Cells["D" + (index + 14) + ":" + "J" + (index + 14)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 14) + ":" + "J" + (index + 14)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 14) + ":" + "J" + (index + 14)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 14) + ":" + "J" + (index + 14)].Style.Font.Size = 9;


                worksheet.Cells["A" + (index + 15) + ":" + "C" + (index + 15)].Merge = true;
                worksheet.Cells["A" + (index + 15) + ":" + "C" + (index + 15)].Value = "Teléfono Móvil";
                worksheet.Cells["A" + (index + 15) + ":" + "C" + (index + 15)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 15) + ":" + "C" + (index + 15)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 15) + ":" + "C" + (index + 15)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 15) + ":" + "C" + (index + 15)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 15) + ":" + "J" + (index + 15)].Merge = true;
                worksheet.Cells["D" + (index + 15) + ":" + "J" + (index + 15)].Value = cotizacion.Prov_movil;
                worksheet.Cells["D" + (index + 15) + ":" + "J" + (index + 15)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 15) + ":" + "J" + (index + 15)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 15) + ":" + "J" + (index + 15)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 15) + ":" + "J" + (index + 15)].Style.Font.Size = 9;

                worksheet.Cells["A" + (index + 16) + ":" + "C" + (index + 16)].Merge = true;
                worksheet.Cells["A" + (index + 16) + ":" + "C" + (index + 16)].Value = "Vigencia de la oferta";
                worksheet.Cells["A" + (index + 16) + ":" + "C" + (index + 16)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 16) + ":" + "C" + (index + 16)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 16) + ":" + "C" + (index + 16)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 16) + ":" + "C" + (index + 16)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 16) + ":" + "J" + (index + 16)].Merge = true;
                worksheet.Cells["D" + (index + 16) + ":" + "J" + (index + 16)].Value = cotizacion.Prov_vigOferta;
                worksheet.Cells["D" + (index + 16) + ":" + "J" + (index + 16)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 16) + ":" + "J" + (index + 16)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 16) + ":" + "J" + (index + 16)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 16) + ":" + "J" + (index + 16)].Style.Font.Size = 9;


                worksheet.Cells["A" + (index + 17) + ":" + "C" + (index + 17)].Merge = true;
                worksheet.Cells["A" + (index + 17) + ":" + "C" + (index + 17)].Value = "Registro Sanitario";
                worksheet.Cells["A" + (index + 17) + ":" + "C" + (index + 17)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + (index + 17) + ":" + "C" + (index + 17)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + (index + 17) + ":" + "C" + (index + 17)].Style.Font.Bold = true;
                worksheet.Cells["A" + (index + 17) + ":" + "C" + (index + 17)].Style.Font.Size = 9;


                worksheet.Cells["D" + (index + 17) + ":" + "J" + (index + 17)].Merge = true;
                worksheet.Cells["D" + (index + 17) + ":" + "J" + (index + 17)].Value = cotizacion.Prog_RegiSanitario;
                worksheet.Cells["D" + (index + 17) + ":" + "J" + (index + 17)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + (index + 17) + ":" + "J" + (index + 17)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["D" + (index + 17) + ":" + "J" + (index + 17)].Style.Font.Bold = true;
                worksheet.Cells["D" + (index + 17) + ":" + "J" + (index + 17)].Style.Font.Size = 9;


                worksheet.Row(index + 24).Height = 27.5;
                worksheet.Cells["C" + (index + 24) + ":" + "F" + (index + 24)].Merge = true;
                worksheet.Cells["C" + (index + 24) + ":" + "F" + (index + 24)].Value = "Firma, nombres y apellidos del proveedor o representante legal o Persona autorizada para emitir cotizaciones";
                worksheet.Cells["C" + (index + 24) + ":" + "F" + (index + 24)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C" + (index + 24) + ":" + "F" + (index + 24)].Style.Font.Size = 9;
                worksheet.Cells["C" + (index + 24) + ":" + "F" + (index + 24)].Style.WrapText = true;



                worksheet.View.ZoomScale = 100;

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
            worksheet.Column(1).Width = 7.29 + 0.71;
            worksheet.Column(2).Width = 9.71 + 0.71;
            worksheet.Column(3).Width = 26.14 + 0.71;
            worksheet.Column(4).Width = 9.57 + 0.71;
            worksheet.Column(5).Width = 8.86 + 0.71;
            worksheet.Column(6).Width = 8.57 + 0.71;
            worksheet.Column(7).Width = 10.71 + 0.71;
            worksheet.Column(8).Width = 10.57 + 0.71;
            worksheet.Column(9).Width = 8.57 + 0.71;
            worksheet.Column(10).Width = 11.57 + 0.71;


            worksheet.Row(1).Height = 26.5;
            worksheet.Row(2).Height = 6.5;
            worksheet.Row(3).Height = 17;
            worksheet.Row(4).Height = 15.5;
            worksheet.Row(5).Height = 5;
            worksheet.Row(6).Height = 12.75;
            worksheet.Row(7).Height = 13.25;
            worksheet.Row(8).Height = 13.25;
            worksheet.Row(9).Height = 13.25; 
            
            worksheet.Row(10).Height = 12.75;
            worksheet.Row(11).Height = 00.00;
            worksheet.Row(12).Height = 12.75;
            worksheet.Row(13).Height = 23.5;
            worksheet.Row(14).Height = 12.5;
            worksheet.Row(15).Height = 29;

            


        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["I1:J1"].Merge = true;
            worksheet.Cells["A3:J3"].Merge = true;
            worksheet.Cells["A4:E4"].Merge = true;
            worksheet.Cells["A7:E7"].Merge = true;
            worksheet.Cells["A9:B9"].Merge = true;
            worksheet.Cells["A10:J13"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A15,B15,C15,D15,E15,F15,G15,H15,I15,J15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {

            worksheet.Cells["A15:J15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }

        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
