using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato22_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato22_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización LoretoF1");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(80, 30);

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["A2"].Value = "ANEXO Nº 05";
                workSheet.Cells["A2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A2"].Style.Font.Bold = true;
                workSheet.Cells["A2"].Style.Font.Size = 10;

                workSheet.Cells["B3"].Value = "FORMATO DE COTIZACIÓN DE BIENES";
                workSheet.Cells["B3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B3"].Style.Font.Bold = true;
                workSheet.Cells["B3"].Style.Font.Size = 10;

                workSheet.Cells["B5"].Value = "Lima " + cotizacion.Fecha_Cotizacion.ToLongDateString();
                workSheet.Cells["B5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B5"].Style.Font.Size = 9;

                workSheet.Cells["B6"].Value = "Señores ";
                workSheet.Cells["B6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B6"].Style.Font.Bold = true;
                workSheet.Cells["B6"].Style.Font.Size = 9;

                workSheet.Cells["B7"].Value = "ESSALUD - Red Loreto ";
                workSheet.Cells["B7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B7"].Style.Font.Bold = true;
                workSheet.Cells["B7"].Style.Font.Size = 10;

                workSheet.Cells["B8"].Value = "De mi consideración";
                workSheet.Cells["B8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B8"].Style.Font.Size = 10;

                workSheet.Cells["B10"].Value = "En respuesta a la solicitud de cotización sobre la 'ADQUISICIÓN ..........................', y después de haber analizado las especificaciones técnicas del mencionado requerimiento," +
                    " las mismas que acepto en todos sus extremos, indico que cumplo con TODOS los requerimientos solicitados.";
                workSheet.Cells["B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B10"].Style.Font.Size = 9;
                workSheet.Cells["B10"].Style.WrapText = true;

                workSheet.Cells["B11"].Value = "Asimismo, declaro que las características técnicas de los bienes cotizados por mi representada se ajustan a lo requerido por su Entidad." +
                    " En tal sentido, indico que el costo total por la Cotización requerida es la que detallo a continuación:";
                workSheet.Cells["B11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B11"].Style.Font.Size = 9;
                workSheet.Cells["B11"].Style.WrapText = true;

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);


                workSheet.Cells["B13"].Value = "Item";
                workSheet.Cells["B13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B13"].Style.Font.Size = 7;
                workSheet.Cells["B13"].Style.WrapText = true;
                workSheet.Cells["B13"].Style.Font.Bold = true;

                workSheet.Cells["C13"].Value = "Codigo de Material";
                workSheet.Cells["C13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C13"].Style.Font.Size = 7;
                workSheet.Cells["C13"].Style.WrapText = true;
                workSheet.Cells["C13"].Style.Font.Bold = true;

                workSheet.Cells["D13"].Value = "Descripcion";
                workSheet.Cells["D13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D13"].Style.Font.Size = 7;
                workSheet.Cells["D13"].Style.WrapText = true;
                workSheet.Cells["D13"].Style.Font.Bold = true;

                workSheet.Cells["E13"].Value = "Cantidad";
                workSheet.Cells["E13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E13"].Style.Font.Size = 7;
                workSheet.Cells["E13"].Style.WrapText = true;
                workSheet.Cells["E13"].Style.Font.Bold = true;

                workSheet.Cells["F13"].Value = "Unidad de Medida";
                workSheet.Cells["F13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F13"].Style.Font.Size = 7;
                workSheet.Cells["F13"].Style.WrapText = true;
                workSheet.Cells["F13"].Style.Font.Bold = true;

                workSheet.Cells["G13"].Value = "Marca";
                workSheet.Cells["G13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G13"].Style.Font.Size = 7;
                workSheet.Cells["G13"].Style.WrapText = true;
                workSheet.Cells["G13"].Style.Font.Bold = true;

                workSheet.Cells["H13"].Value = "Modelo";
                workSheet.Cells["H13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H13"].Style.Font.Size = 7;
                workSheet.Cells["H13"].Style.WrapText = true;
                workSheet.Cells["H13"].Style.Font.Bold = true;


                workSheet.Cells["I13"].Value = "Pais de Procedencia";
                workSheet.Cells["I13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I13"].Style.Font.Size = 7;
                workSheet.Cells["I13"].Style.WrapText = true;
                workSheet.Cells["I13"].Style.Font.Bold = true;


                workSheet.Cells["J13"].Value = "PRECIO UNITARIO (Soles) INCLUIDO IGV";
                workSheet.Cells["J13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J13"].Style.Font.Size = 7;
                workSheet.Cells["J13"].Style.WrapText = true;
                workSheet.Cells["J13"].Style.Font.Bold = true;

                workSheet.Cells["K13"].Value = "PRECIO TOTAL (Soles) INCLUIDO IGV";
                workSheet.Cells["K13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K13"].Style.Font.Size = 7;
                workSheet.Cells["K13"].Style.WrapText = true;
                workSheet.Cells["K13"].Style.Font.Bold = true;

                int index = 16;
                foreach (Coti_Formato22_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Row(index).Height = 32.25;
                    workSheet.Cells["B" + index].Value = index - 15;
                    workSheet.Cells["B" + index].Style.Font.Size = 7;
                    workSheet.Cells["B" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;
                    workSheet.Cells["B" + index].Style.Numberformat.Format = "0";

                    
                    workSheet.Cells["C" + index].Value = item.CodMaterial;
                    workSheet.Cells["C" + index].Style.Font.Size = 7;
                    workSheet.Cells["C" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;


                    workSheet.Cells["D" + index].Value = item.Descripcion;
                    workSheet.Cells["D" + index].Style.Font.Size = 7;
                    workSheet.Cells["D" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;

                    workSheet.Cells["E" + index].Value = item.Cantidad;
                    workSheet.Cells["E" + index].Style.Font.Size = 7;
                    workSheet.Cells["E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;
                    workSheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["F" + index].Value = item.UndMedida;
                    workSheet.Cells["F" + index].Style.Font.Size = 7;
                    workSheet.Cells["F" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;

                    workSheet.Cells["G" + index].Value = item.Marca;
                    workSheet.Cells["G" + index].Style.Font.Size = 7;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;

                    workSheet.Cells["H" + index].Value = item.Modelo;
                    workSheet.Cells["H" + index].Style.Font.Size = 7;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;

                    workSheet.Cells["I" + index].Value = item.Procedencia;
                    workSheet.Cells["I" + index].Style.Font.Size = 7;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Cells["J" + index].Value = item.PreUnitario;
                    workSheet.Cells["J" + index].Style.Font.Size = 7;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;
                    workSheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["K" + index].Value = item.PreUnitario;
                    workSheet.Cells["K" + index].Style.Font.Size = 7;
                    workSheet.Cells["K" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;
                    workSheet.Cells["K" + index].Style.Numberformat.Format = "#,##0.00";


                    index++;
                }

                workSheet.Row(index).Height = 18.00;
                workSheet.Cells["B" + index + ":I" + index].Merge = true;
                workSheet.Cells["B" + index + ":I" + index].Value = "VALOR TOTAL DE LA COTIZACIÓN";
                workSheet.Cells["B" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":I" + index].Style.Font.Size = 9;
                workSheet.Cells["B" + index + ":I" + index].Style.Font.Name = "Arial";

                workSheet.Cells["J" + index + ":K" + index].Merge = true;
                workSheet.Cells["J" + index + ":K" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["J" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["J" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["J" + index + ":K" + index].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["J" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["J" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["J" + index + ":K" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 59.25;
                workSheet.Row(index+1).Height = 15.00;
                workSheet.Cells["B" + index + ":K" + (index + 1)].Merge = true;
                workSheet.Cells["B" + index + ":K" + (index + 1)].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso, los costos laborales conforme la legislación vigente," +
                    " así como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y/o servicio a contratar; excepto la de aquellos proveedores que gocen de alguna exoneración legal, no incluirán en el precio de su oferta los tributos respectivos. Asimismo, declaro bajo juramente que, mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, conforme lo establece el artículo 11 del " +
                    "Texto Único Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado, aprobado por Decreto Supremo N° 082-2019-EF.";
                workSheet.Cells["B" + index + ":K" + (index + 1)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":K" + (index + 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["B" + index + ":K" + (index + 1)].Style.Font.Size = 8;
                workSheet.Cells["B" + index + ":K" + (index + 1)].Style.Font.Name = "Arial";

                index++;
                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Razón social";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Nº RUC";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Plazo de entrega";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Forma de pago";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_FormaPago;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Garantía";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_Garantia;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Correo electrónico";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_Email;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Teléfono Fijo";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_ContTelefono;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Persona de contacto";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Teléfono móvil";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_ContCelular;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 15.00;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "Vigencia de oferta";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 11;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Arial";

                workSheet.Cells["E" + index + ":K" + index].Merge = true;
                workSheet.Cells["E" + index + ":K" + index].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["E" + index + ":K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index + ":K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Size = 11;
                workSheet.Cells["E" + index + ":K" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index+8).Height = 15.00;
                workSheet.Cells["B" + (index + 8) + ":K" + (index + 8)].Merge = true;
                workSheet.Cells["B" + (index + 8) + ":K" + (index + 8)].Value = "Firma, nombres y apellidos del proveedor o representante legal o ";
                workSheet.Cells["B" + (index + 8) + ":K" + (index + 8)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + (index + 8) + ":K" + (index + 8)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + (index + 8) + ":K" + (index + 8)].Style.Font.Size = 9;
                workSheet.Cells["B" + (index + 8) + ":K" + (index + 8)].Style.Font.Name = "Arial";

                workSheet.Row(index + 9).Height = 15.00;
                workSheet.Cells["B" + (index + 9) + ":K" + (index + 9)].Merge = true;
                workSheet.Cells["B" + (index + 9) + ":K" + (index + 9)].Value = "persona autorizada para emitir cotizaciones";
                workSheet.Cells["B" + (index + 9) + ":K" + (index + 9)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + (index + 9) + ":K" + (index + 9)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + (index + 9) + ":K" + (index + 9)].Style.Font.Size = 9;
                workSheet.Cells["B" + (index + 9) + ":K" + (index + 9)].Style.Font.Name = "Arial";

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
            workSheet.Column(1).Width = 0.75 + 0.71;
            workSheet.Column(2).Width = 4 + 0.71;
            workSheet.Column(3).Width =  7.43+ 0.71;
            workSheet.Column(4).Width = 16.57 + 0.71;
            workSheet.Column(5).Width = 10.71 + 0.71;
            workSheet.Column(6).Width = 8.00 + 0.71;
            workSheet.Column(7).Width = 10.14 + 0.71;
            workSheet.Column(8).Width = 6.57 + 0.71;
            workSheet.Column(9).Width = 10.71  + 0.71;
            workSheet.Column(10).Width = 10.71 + 0.71;
            workSheet.Column(11).Width = 9.71  + 0.71;
            workSheet.Column(12).Width = 0.67 + 0.71;

            workSheet.Row(1).Height = 12.00;
            workSheet.Row(2).Height = 12.75;
            workSheet.Row(3).Height = 18.00;
            workSheet.Row(4).Height = 12.00;
            workSheet.Row(5).Height = 12.00;
            workSheet.Row(6).Height = 12.00;
            workSheet.Row(7).Height = 12.75;
            workSheet.Row(8).Height = 12.00;
            workSheet.Row(9).Height = 12.00; 
            workSheet.Row(10).Height = 31.5;
            workSheet.Row(11).Height = 32.25;

            workSheet.Row(12).Height = 12.00;
            workSheet.Row(13).Height = 22.5;
            workSheet.Row(14).Height = 16.5;
            workSheet.Row(15).Height = 21.75;
        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {

            workSheet.Cells["A2:K2"].Merge = true;
            workSheet.Cells["B3:K3"].Merge = true;
            workSheet.Cells["B5:D5"].Merge = true;
            workSheet.Cells["B6:C6"].Merge = true;
            workSheet.Cells["B7:D7"].Merge = true;
            workSheet.Cells["B8:D8"].Merge = true;

            workSheet.Cells["B10:K10"].Merge = true;
            workSheet.Cells["B11:K11"].Merge = true;


            workSheet.Cells["B13:B15"].Merge = true;
            workSheet.Cells["C13:C15"].Merge = true;
            workSheet.Cells["D13:D15"].Merge = true;
            workSheet.Cells["E13:E15"].Merge = true;
            workSheet.Cells["F13:F15"].Merge = true;
            workSheet.Cells["G13:G15"].Merge = true;
            workSheet.Cells["H13:H15"].Merge = true;
            workSheet.Cells["I13:I15"].Merge = true;
            workSheet.Cells["J13:J15"].Merge = true;
            workSheet.Cells["K13:K15"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["B3:K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["B13:B15"].Style.Border.BorderAround(ExcelBorderStyle.Double);
            workSheet.Cells["C13:C15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D13:D15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E13:E15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F13:F15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G13:G15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H13:H15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I13:I15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J13:J15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K13:K15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
           
        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["B13:K15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }



     
    }
}
