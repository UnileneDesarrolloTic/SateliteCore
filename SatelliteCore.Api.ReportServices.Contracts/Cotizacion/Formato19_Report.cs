using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato19_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato19_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Junin");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 50);

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["A2"].Value = "ANEXO Nº 05";
                workSheet.Cells["A2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A2"].Style.Font.Bold = true;

                workSheet.Cells["A3"].Value = "FORMATO DE COTIZACION DE BIENES";
                workSheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A3"].Style.Font.Bold = true;
                workSheet.Cells["A3"].Style.Font.Size = 14;


                workSheet.Cells["A5"].Value = "Cotización Nro "+ cotizacion.Prov_NroCotizacion+"-2022";
                workSheet.Cells["A5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A5"].Style.Font.Bold = true;
                workSheet.Cells["A5"].Style.Font.Size = 11;
                workSheet.Cells["A5"].Style.Font.UnderLine = true;

                workSheet.Cells["A7"].Value = "Lima " + cotizacion.Fecha_Cotizacion.ToLongDateString();
                workSheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A7"].Style.Font.Size = 11;
                workSheet.Cells["A7"].Style.Font.Name = "Calibri";

                workSheet.Cells["A9"].Value = "Señores: ";
                workSheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A9"].Style.Font.Size = 11;
               


                workSheet.Cells["A10"].Value = "RED ASISTENCIAL ESSALUD: JUNIN";
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A10"].Style.Font.Size = 11;

                workSheet.Cells["A12"].Value = "De mi consideración:";
                workSheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A12"].Style.Font.Size = 11;


                workSheet.Cells["A14"].Value = "En respuesta a la solicitud de cotización sobre Dispositivos médicos   " +
                    "y después de haber analizado las especificaciones técnicas del mencionado requerimiento, las mismas que acepto en todos sus extremos, indico que cumplo con TODOS los requerimientos solicitados."+
                    "Asimismo, declaro que las características técnicas de los bienes cotizados por mi representada se ajustan a lo requerido por su Entidad. En tal sentido, indico que el costo total por la solución requerida es la que detallo a continuación";
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A14"].Style.Font.Size = 10;
                workSheet.Cells["A14"].Style.WrapText = true;

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);

                //DETALLE 

                workSheet.Cells["A16"].Value = "Ítem";
                workSheet.Cells["A16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A16"].Style.Font.Size = 8;
                workSheet.Cells["A16"].Style.WrapText = true;
                workSheet.Cells["A16"].Style.Font.Bold = true;

                workSheet.Cells["B16"].Value = "Descripción";
                workSheet.Cells["B16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B16"].Style.Font.Size = 8;
                workSheet.Cells["B16"].Style.WrapText = true;
                workSheet.Cells["B16"].Style.Font.Bold = true;

                workSheet.Cells["C16"].Value = "Cantidad";
                workSheet.Cells["C16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C16"].Style.Font.Size = 8;
                workSheet.Cells["C16"].Style.WrapText = true;
                workSheet.Cells["C16"].Style.Font.Bold = true;

                workSheet.Cells["D16"].Value = "Unidad de Medida";
                workSheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D16"].Style.Font.Size = 8;
                workSheet.Cells["D16"].Style.WrapText = true;
                workSheet.Cells["D16"].Style.Font.Bold = true;


                workSheet.Cells["E16"].Value = "Marca";
                workSheet.Cells["E16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E16"].Style.Font.Size = 8;
                workSheet.Cells["E16"].Style.WrapText = true;
                workSheet.Cells["E16"].Style.Font.Bold = true;


                workSheet.Cells["F16"].Value = "Modelo";
                workSheet.Cells["F16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F16"].Style.Font.Size = 8;
                workSheet.Cells["F16"].Style.WrapText = true;
                workSheet.Cells["F16"].Style.Font.Bold = true;

                workSheet.Cells["G16"].Value = "País de Procedencia";
                workSheet.Cells["G16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G16"].Style.Font.Size = 8;
                workSheet.Cells["G16"].Style.WrapText = true;
                workSheet.Cells["G16"].Style.Font.Bold = true;



                workSheet.Cells["H16"].Value = "Registro Sanitario  Si / No";
                workSheet.Cells["H16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H16"].Style.Font.Size = 8;
                workSheet.Cells["H16"].Style.WrapText = true;
                workSheet.Cells["H16"].Style.Font.Bold = true;


                workSheet.Cells["I16"].Value = "PRECIO UNITARIO INC. IGV (soles)";
                workSheet.Cells["I16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I16"].Style.Font.Size = 8;
                workSheet.Cells["I16"].Style.WrapText = true;
                workSheet.Cells["I16"].Style.Font.Bold = true;



                workSheet.Cells["J16"].Value = "PRECIO TOTAL INC. IGV (soles)";
                workSheet.Cells["J16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J16"].Style.Font.Size = 8;
                workSheet.Cells["J16"].Style.WrapText = true;
                workSheet.Cells["J16"].Style.Font.Bold = true;

                int index = 18;
                foreach (Coti_Formato19_Detalle item in cotizacion.Detalle)
                {

                    workSheet.Row(index).Height = 80.25;
                    workSheet.Cells["A" + index].Value = index - 17;
                    workSheet.Cells["A" + index].Style.Font.Size = 9;
                    workSheet.Cells["A" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.WrapText = true;
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                    workSheet.Cells["B" + index].Value = item.Descripcion;
                    workSheet.Cells["B" + index].Style.Font.Size = 9;
                    workSheet.Cells["B" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;

                    workSheet.Cells["C" + index].Value = item.Cantidad;
                    workSheet.Cells["C" + index].Style.Font.Size = 9;
                    workSheet.Cells["C" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["D" + index].Value = item.UndMedida;
                    workSheet.Cells["D" + index].Style.Font.Size = 9;
                    workSheet.Cells["D" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = item.Marca;
                    workSheet.Cells["E" + index].Style.Font.Size = 9;
                    workSheet.Cells["E" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = item.Modelo;
                    workSheet.Cells["F" + index].Style.Font.Size = 9;
                    workSheet.Cells["F" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = item.Procedencia;
                    workSheet.Cells["G" + index].Style.Font.Size = 9;
                    workSheet.Cells["G" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = item.RegSanitario;
                    workSheet.Cells["H" + index].Style.Font.Size = 9;
                    workSheet.Cells["H" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = item.PreUnitario;
                    workSheet.Cells["I" + index].Style.Font.Size = 9;
                    workSheet.Cells["I" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["J" + index].Value = item.PreTotal;
                    workSheet.Cells["J" + index].Style.Font.Size = 9;
                    workSheet.Cells["J" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";

                    index++;

                }

                workSheet.Row(index).Height = 18.00;
                workSheet.Cells["A" + index + ":I" + index].Merge = true;
                workSheet.Cells["A" + index + ":I" + index].Value = "VALOR TOTAL DE LA COTIZACIÓN";
                workSheet.Cells["A" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":I" + index].Style.Font.Size = 9;
                workSheet.Cells["A" + index + ":I" + index].Style.Font.Name = "Arial";

                workSheet.Cells["J" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["J" + index ].Style.Font.Bold = true;
                workSheet.Cells["J" + index].Style.Font.Size = 11;
                workSheet.Cells["J" + index].Style.Font.Name = "Arial";

                index++;
                    workSheet.Row(index).Height = 12;

                index++;
                workSheet.Row(index).Height = 65.75;
                workSheet.Cells["A" + index + ":J" + index].Merge = true;
                workSheet.Cells["A" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":J" + index].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso," +
                    " los costos laborales conforme la legislación vigente, así como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y/o servicio a contratar, excepto la de aquellos proveedores que gocen de alguna exoneración legal, " +
                    "no incluirán en el precio de su oferta los tributos respectivos. Asimismo, declaro bajo juramento que, mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, conforme lo" +
                    " establece el articulo 11 del Texto Único Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado, aprobado por Decreto Supremo N° 082-2019-EF.";
                workSheet.Cells["A" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Razón social";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "N° RUC";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Plazo de entrega";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;



                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Forma de pago";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;



                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_FormaPago;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Garantía";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Garantia;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;


                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Garantía";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Garantia+ "Vigencia Producto : "+cotizacion.Prov_vigProducto;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;


                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Correo electrónico";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;


                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Email;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Teléfono fijo";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Persona de contacto";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;


                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Teléfono móvil";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_Celular;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Vigencia de oferta";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Value = "Registro Sanitario";
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Size = 9;

                workSheet.Cells["C" + index + ":J" + index].Merge = true;
                workSheet.Cells["C" + index + ":J" + index].Value = cotizacion.Prov_RegSanitario;
                workSheet.Cells["C" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":J" + index].Style.Font.Size = 8;

                index++;
                workSheet.Row(index).Height = 15;
                index++;
                workSheet.Row(index).Height = 15;
                index++;
                workSheet.Row(index).Height = 21.75;
                index++;
                workSheet.Row(index).Height = 19.5;
                index++;
                workSheet.Row(index).Height = 15;
                workSheet.Cells["B" + index + ":H" + index].Merge = true;
                workSheet.Cells["B" + index + ":H" + index].Value = "Firma, nombre y apellido del proveedoro representante legal o";
                workSheet.Cells["B" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                index++;
                workSheet.Row(index).Height = 15;
                workSheet.Cells["B" + index + ":H" + index].Merge = true;
                workSheet.Cells["B" + index + ":H" + index].Value = "persona autorizada para emitir cotizaciones";
                workSheet.Cells["B" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                TextoNegrita(workSheet);

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

            workSheet.Column(1).Width = 7.71 + 0.71;
            workSheet.Column(2).Width = 18.57 + 0.71;
            workSheet.Column(3).Width = 9.00 + 0.71;
            workSheet.Column(4).Width = 5.86 + 0.71;
            workSheet.Column(5).Width = 9.00 + 0.71;
            workSheet.Column(6).Width = 9.57 + 0.71;
            workSheet.Column(7).Width = 10.00 + 0.71;
            workSheet.Column(8).Width = 7.57 + 0.71;
            workSheet.Column(9).Width = 11.00 + 0.71;
            workSheet.Column(10).Width = 13.00 + 0.71;

            workSheet.Row(1).Height = 39.75;
            workSheet.Row(2).Height = 15.00;
            workSheet.Row(3).Height = 24.00;
            workSheet.Row(4).Height = 21.00;
            workSheet.Row(5).Height = 17.25;
            workSheet.Row(6).Height = 13.5;
            workSheet.Row(7).Height = 15.00;
            workSheet.Row(8).Height = 7.5;
            workSheet.Row(9).Height = 15.00;
            workSheet.Row(10).Height = 15.00;
            workSheet.Row(11).Height = 12.00;
            workSheet.Row(12).Height = 15.00;
            workSheet.Row(13).Height = 4.5;
            workSheet.Row(14).Height = 65.25;
            workSheet.Row(15).Height = 5.25;
            workSheet.Row(16).Height = 21.25;
            workSheet.Row(17).Height = 21.75;

        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A1:J1"].Merge = true;
            workSheet.Cells["A2:J2"].Merge = true;
            workSheet.Cells["A3:J3"].Merge = true;
            workSheet.Cells["A4:J4"].Merge = true;

            workSheet.Cells["A5:C5"].Merge = true;
            workSheet.Cells["A7:C7"].Merge = true;

            workSheet.Cells["A9:B9"].Merge = true;
            workSheet.Cells["A10:C10"].Merge = true;
            workSheet.Cells["A12:C12"].Merge = true;
            workSheet.Cells["A14:J14"].Merge = true;

            workSheet.Cells["A16:A17"].Merge = true;
            workSheet.Cells["B16:B17"].Merge = true;
            workSheet.Cells["C16:C17"].Merge = true;
            workSheet.Cells["D16:D17"].Merge = true;
            workSheet.Cells["E16:E17"].Merge = true;
            workSheet.Cells["F16:F17"].Merge = true;
            workSheet.Cells["G16:G17"].Merge = true;
            workSheet.Cells["H16:H17"].Merge = true;
            workSheet.Cells["I16:I17"].Merge = true;
            workSheet.Cells["J16:J17"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A3:J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A16:A17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B16:B17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C16:C17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D16:D17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E16:E17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F16:F17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G16:G17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H16:H17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I16:I17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J16:J17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
         
            workSheet.Cells["A16:J17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
        private static void TextoNegrita(ExcelWorksheet workSheet)
        {

        }

    }
}
