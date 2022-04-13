using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato13_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato13_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Sabogal");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 50);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                workSheet.Cells["A5"].Value = "ANEXO N° 05";
                workSheet.Cells["A5"].Style.Font.Size = 12;
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A5"].Style.Font.Bold = true;

                workSheet.Cells["A6"].Value = "FORMATO DE COTIZACION DE BIENES";
                workSheet.Cells["A6"].Style.Font.Size = 12;
                workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A6"].Style.Font.Bold = true;

                workSheet.Cells["A7"].Value ="Lima, " + cotizacion.Fecha_Cotizacion.ToLongDateString();
                workSheet.Cells["A7"].Style.Font.Size = 12;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A7"].Style.Font.Bold = true;
               


                workSheet.Cells["A8"].Value = "Señores: ";
                workSheet.Cells["A8"].Style.Font.Size = 11;
                workSheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A8"].Style.Font.Bold = true;

                workSheet.Cells["K8"].Value = "Cotiz. Nro " + cotizacion.Prov_NroCotizacion + " Unilene S.A.C";
                workSheet.Cells["K8"].Style.Font.Size = 11;
                workSheet.Cells["K8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K8"].Style.Font.Bold = true;
                workSheet.Cells["K8"].Style.Font.UnderLine = true;

                workSheet.Cells["A9"].Value = "ESSALUD CUSCO";
                workSheet.Cells["A9"].Style.Font.Size = 11;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A9"].Style.Font.Bold = true;

                workSheet.Cells["A10"].Value = "De mi consideracion:";
                workSheet.Cells["A10"].Style.Font.Size = 11;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A10"].Style.Font.Bold = true;

                workSheet.Cells["K10"].Value = "COT. N° " + cotizacion.Prov_NroCotizacion;
                workSheet.Cells["K10"].Style.Font.Size = 11;
                workSheet.Cells["K10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K10"].Style.Font.Bold = true;
                workSheet.Cells["K10"].Style.Font.UnderLine = true;

                workSheet.Cells["A11"].Value = "En respuesta a la solictud de cotizacion sobre la " + "'ADQUISICION DE MATERIAL MEDICO'," + " y despues de haber analizado las especificaciones tecnicas del mencionado" +
                    " requerimiento, las mismas que acepto en todos sus extremos, indico que cumplo con TODOS los requerimientos solicitados.";
                workSheet.Cells["A11"].Style.Font.Size = 11;
                workSheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A11"].Style.WrapText = true;
                workSheet.Cells["A11"].Style.Font.Bold = true;


                workSheet.Cells["A12"].Value = "Asimismo, declaro que las caracteristicas tecnicas de los bienes cotizados por mi representada se ajustan a lo requerido por su Entidad. En tal sentido, indico que " +
                    "el costo total por la solucion requerida es la que detallo a continuacion.";
                workSheet.Cells["A12"].Style.Font.Size = 11;
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A12"].Style.WrapText = true;
                workSheet.Cells["A12"].Style.Font.Bold = true;


                workSheet.Cells["A12"].Value = "Asimismo, declaro que las caracteristicas tecnicas de los bienes cotizados por mi representada se ajustan a lo requerido por su Entidad. En tal sentido, indico que " +
                   "el costo total por la solucion requerida es la que detallo a continuacion.";
                workSheet.Cells["A12"].Style.Font.Size = 11;
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A12"].Style.WrapText = true;
                workSheet.Cells["A12"].Style.Font.Bold = true;


                workSheet.Cells["A13:H13"].Value = "OPO-OGD-GRACU-ESSALUD-2022";
                workSheet.Cells["A13:H13"].Style.Font.Size = 11;
                workSheet.Cells["A13:H13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A13:H13"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A13:H13"].Style.WrapText = true;
                workSheet.Cells["A13:H13"].Style.Font.Bold = true;
                workSheet.Cells["A13:H13"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FF3A00"));





                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);


                workSheet.Cells["A14"].Value = "N° ITEM";
                workSheet.Cells["A14"].Style.Font.Size = 10;
                workSheet.Cells["A14"].Style.Font.Name = "Arial";
                workSheet.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A14"].Style.WrapText = true;
                workSheet.Cells["A14"].Style.Font.Bold = true;
                workSheet.Cells["A14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["B14"].Value = "CODIGO SAP";
                workSheet.Cells["B14"].Style.Font.Size = 10;
                workSheet.Cells["B14"].Style.Font.Name = "Arial";
                workSheet.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B14"].Style.WrapText = true;
                workSheet.Cells["B14"].Style.Font.Bold = true;
                workSheet.Cells["B14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));


                workSheet.Cells["C14"].Value = "DESCRIPCION DEL ITEM";
                workSheet.Cells["C14"].Style.Font.Size = 10;
                workSheet.Cells["C14"].Style.Font.Name = "Arial";
                workSheet.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C14"].Style.WrapText = true;
                workSheet.Cells["C14"].Style.Font.Bold = true;
                workSheet.Cells["C14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["D14"].Value = "CANTIDAD TOTAL";
                workSheet.Cells["D14"].Style.Font.Size = 10;
                workSheet.Cells["D14"].Style.Font.Name = "Arial";
                workSheet.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D14"].Style.WrapText = true;
                workSheet.Cells["D14"].Style.Font.Bold = true;
                workSheet.Cells["D14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));


                workSheet.Cells["E14"].Value = "U.M";
                workSheet.Cells["E14"].Style.Font.Size = 10;
                workSheet.Cells["E14"].Style.Font.Name = "Arial";
                workSheet.Cells["E14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E14"].Style.WrapText = true;
                workSheet.Cells["E14"].Style.Font.Bold = true;
                workSheet.Cells["E14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["F14"].Value = "PLAZO DE ENTREGA";
                workSheet.Cells["F14"].Style.Font.Size = 10;
                workSheet.Cells["F14"].Style.Font.Name = "Arial";
                workSheet.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F14"].Style.WrapText = true;
                workSheet.Cells["F14"].Style.Font.Bold = true;
                workSheet.Cells["F14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));


                workSheet.Cells["G14"].Value = "PROCEDENCIA";
                workSheet.Cells["G14"].Style.Font.Size = 10;
                workSheet.Cells["G14"].Style.Font.Name = "Arial";
                workSheet.Cells["G14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G14"].Style.WrapText = true;
                workSheet.Cells["G14"].Style.Font.Bold = true;
                workSheet.Cells["G14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["H14"].Value = "PRESENTACION";
                workSheet.Cells["H14"].Style.Font.Size = 10;
                workSheet.Cells["H14"].Style.Font.Name = "Arial";
                workSheet.Cells["H14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H14"].Style.WrapText = true;
                workSheet.Cells["H14"].Style.Font.Bold = true;
                workSheet.Cells["H14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["I14"].Value = "MARCA/ MODELO";
                workSheet.Cells["I14"].Style.Font.Size = 10;
                workSheet.Cells["I14"].Style.Font.Name = "Arial";
                workSheet.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I14"].Style.WrapText = true;
                workSheet.Cells["I14"].Style.Font.Bold = true;
                workSheet.Cells["I14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));


                workSheet.Cells["J14"].Value = "CUMPLE AL 100% CON LAS EE. TT.";
                workSheet.Cells["J14"].Style.Font.Size = 10;
                workSheet.Cells["J14"].Style.Font.Name = "Arial";
                workSheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J14"].Style.WrapText = true;
                workSheet.Cells["J14"].Style.Font.Bold = true;
                workSheet.Cells["J14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["K14"].Value = "VIGENCIA DE LA OFERTA";
                workSheet.Cells["K14"].Style.Font.Size = 10;
                workSheet.Cells["K14"].Style.Font.Name = "Arial";
                workSheet.Cells["K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K14"].Style.WrapText = true;
                workSheet.Cells["K14"].Style.Font.Bold = true;
                workSheet.Cells["K14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["L14"].Value = "RNP";
                workSheet.Cells["L14"].Style.Font.Size = 10;
                workSheet.Cells["L14"].Style.Font.Name = "Arial";
                workSheet.Cells["L14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L14"].Style.WrapText = true;
                workSheet.Cells["L14"].Style.Font.Bold = true;
                workSheet.Cells["L14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));


                workSheet.Cells["M14"].Value = "PRECIO UNITARIO INC. IGV (S/.)";
                workSheet.Cells["M14"].Style.Font.Size = 10;
                workSheet.Cells["M14"].Style.Font.Name = "Arial";
                workSheet.Cells["M14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M14"].Style.WrapText = true;
                workSheet.Cells["M14"].Style.Font.Bold = true;
                workSheet.Cells["M14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                workSheet.Cells["N14"].Value = "PRECIO UNITARIO INC. IGV (S/.)";
                workSheet.Cells["N14"].Style.Font.Size = 10;
                workSheet.Cells["N14"].Style.Font.Name = "Arial";
                workSheet.Cells["N14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N14"].Style.WrapText = true;
                workSheet.Cells["N14"].Style.Font.Bold = true;
                workSheet.Cells["N14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));


                int index = 15;

                foreach (Coti_Formato13_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Cells["A" + index].Value = index - 14;
                    workSheet.Cells["A" + index].Style.Font.Size = 11;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["A" + index].Style.WrapText = true;


                    workSheet.Cells["B" + index].Value = item.CodigoSap;
                    workSheet.Cells["B" + index].Style.Font.Size = 8;
                    workSheet.Cells["B" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;


                    workSheet.Cells["C" + index].Value = item.Descripcion;
                    workSheet.Cells["C" + index].Style.Font.Size = 11;
                    workSheet.Cells["C" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;

                    workSheet.Cells["D" + index].Value = item.Cantidad;
                    workSheet.Cells["D" + index].Style.Font.Size = 11;
                    workSheet.Cells["D" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;
                    workSheet.Cells["D" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["E" + index].Value = item.Um;
                    workSheet.Cells["E" + index].Style.Font.Size = 10;
                    workSheet.Cells["E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;
                    workSheet.Cells["E" + index].Style.Font.Bold = true;

                    workSheet.Cells["F" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["F" + index].Style.Font.Size = 8;
                    workSheet.Cells["F" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;
                    workSheet.Cells["F" + index].Style.Font.Bold = true;

                    workSheet.Cells["G" + index].Value = item.Procedencia;
                    workSheet.Cells["G" + index].Style.Font.Size = 10;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;
                    workSheet.Cells["G" + index].Style.Font.Bold = true;

                    workSheet.Cells["H" + index].Value = item.Presentacion;
                    workSheet.Cells["H" + index].Style.Font.Size = 10;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;
                    workSheet.Cells["H" + index].Style.Font.Bold = true;

                    workSheet.Cells["I" + index].Value = item.MarcaModelo;
                    workSheet.Cells["I" + index].Style.Font.Size = 10;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;
                    workSheet.Cells["I" + index].Style.Font.Bold = true;


                    workSheet.Cells["J" + index].Value = item.CumpleEETT;
                    workSheet.Cells["J" + index].Style.Font.Size = 10;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;
                    workSheet.Cells["J" + index].Style.Font.Bold = true;

                    workSheet.Cells["K" + index].Value = item.VigOferta;
                    workSheet.Cells["K" + index].Style.Font.Size = 9;
                    workSheet.Cells["K" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;
                    workSheet.Cells["K" + index].Style.Font.Bold = true;


                    workSheet.Cells["L" + index].Value = item.Rnp;
                    workSheet.Cells["L" + index].Style.Font.Size = 10;
                    workSheet.Cells["L" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;
                    workSheet.Cells["L" + index].Style.Font.Bold = true;


                    workSheet.Cells["M" + index].Value = item.PreUnitario;
                    workSheet.Cells["M" + index].Style.Font.Size = 10;
                    workSheet.Cells["M" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;
                    workSheet.Cells["M" + index].Style.Font.Bold = true;
                    workSheet.Cells["M" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["N" + index].Value = item.PreTotal;
                    workSheet.Cells["N" + index].Style.Font.Size = 10;
                    workSheet.Cells["N" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;
                    workSheet.Cells["N" + index].Style.Font.Bold = true;
                    workSheet.Cells["N" + index].Style.Numberformat.Format = "#,##0.00";


                    index++;

                }

                workSheet.Cells["A" + index + ":L" + index].Merge = true;
                workSheet.Cells["A" + index + ":L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["M" + index + ":N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["M" + index].Value = "TOTAL S/";
                workSheet.Cells["N" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["N" + index].Style.Numberformat.Format = "#,##0.00";

                index++;

                workSheet.Row(index).Height = 65.5;
                workSheet.Cells["A" + index + ":N" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, inpecciones pruebas y, de ser el caso los costos laborales conforme la legislacion vigente, asi como cualquier" +
                    " otro concepto que pueda tener incidencia sobre el costo del bien y/o servicio a contratar, excepto la de aquellos proveedores que gocen de alguna exoneracion legal," +
                    " no incluiran en el precio de su oferta los tributos respectivos. Asimismo, declaro bajo juramento que, mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, conforme lo establece el articulo 11  " +
                    "del Texto Unico Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado aprobado por Decreto Supremo 082-2019-EF";

                workSheet.Cells["A" + index + ":N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":N" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":N" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":N" + index].Style.JustifyLastLine = true;
                workSheet.Cells["A" + index + ":N" + index].Style.Font.Name = "Calibri";

                index++;
                index++;

                workSheet.Row(index).Height = 26.25;
                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "RAZON SOCIAL PROVEEDOR (P. NATURAL O JURIDICA)";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I" + index + ":J" + index].Merge = true;
                workSheet.Cells["I" + index + ":J" + index].Value = "VIGENCIA - OFERTA " + cotizacion.Prov_VigOferta;
                workSheet.Cells["I" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["I" + index + ":J" + index].Style.Font.Size = 10;
                workSheet.Cells["I" + index + ":J" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["I" + index + ":J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                //firma 
                workSheet.Cells["K" + index + ":N" + (index + 5)].Merge = true;
                workSheet.Cells["K" + index + ":N" + (index + 5)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Merge = true;
                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Value = "Firma, nombre y apellidos del proveedor y representante legal o persona autorizada para emitir cotizaciones";
                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Style.WrapText = true;
                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Style.Font.Size = 9;
                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Style.Font.Name = "Calibri";
                workSheet.Cells["K" + (index + 6) + ":N" + (index + 6)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;

                workSheet.Row(index).Height = 26.25;
                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "RUC";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;

                workSheet.Row(index).Height = 26.25;
                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "FORMA DE PAGO";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_FormaPago;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;

                workSheet.Row(index).Height = 26.25;
                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "CORREO ELECTRONICO";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Email;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;

                workSheet.Row(index).Height = 26.25;
                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "TELEFONO FIJO Y MOVIL";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Telefono+ "-"+cotizacion.Prov_Celular;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;

                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "PERSONA DE CONTACTO";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;

                workSheet.Row(index).Height = 26.25;
                workSheet.Cells["B" + index + ":C" + index].Merge = true;
                workSheet.Cells["B" + index + ":C" + index].Value = "VIGENCIA DE PRODUCTO";
                workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D" + index + ":G" + index].Merge = true;
                workSheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_VigProducto;
                workSheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Size = 10;
                workSheet.Cells["D" + index + ":G" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


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
            workSheet.Column(1).Width = 5.57 + 0.71;
            workSheet.Column(2).Width = 7.57 + 0.71;
            workSheet.Column(3).Width = 25.43 + 0.71;
            workSheet.Column(4).Width = 8.86 + 0.71;
            workSheet.Column(5).Width = 3.57 + 0.71;
            workSheet.Column(6).Width = 10.86 + 0.71;
            workSheet.Column(7).Width = 11.00 + 0.71;
            workSheet.Column(8).Width = 12.14 + 0.71;
            workSheet.Column(9).Width = 10.86 + 0.71;
            workSheet.Column(10).Width = 9.71 + 0.71;
            workSheet.Column(11).Width = 14.00 + 0.71;
            workSheet.Column(12).Width = 5.14 + 0.71;
            workSheet.Column(13).Width = 10.57 + 0.71;
            workSheet.Column(14).Width = 12.00 + 0.71;
            workSheet.Column(15).Width = 10.71 + 0.71;

            workSheet.Row(5).Height = 15.75;
            workSheet.Row(6).Height = 18.75;
            workSheet.Row(7).Height = 15.00;
            workSheet.Row(8).Height = 24.75;
            workSheet.Row(10).Height = 18.75;
            workSheet.Row(11).Height = 36.75;
            workSheet.Row(12).Height = 41.25;
            workSheet.Row(13).Height = 35.25;
            workSheet.Row(14).Height = 55.5;

        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A5:N5"].Merge = true;
            workSheet.Cells["A6:N6"].Merge = true;
            workSheet.Cells["A7:C7"].Merge = true;
            workSheet.Cells["A8:J8"].Merge = true;
            workSheet.Cells["K8:N8"].Merge = true;
            workSheet.Cells["A9:C9"].Merge = true;
            workSheet.Cells["A10:C10"].Merge = true;
            workSheet.Cells["K10:N10"].Merge = true;
            workSheet.Cells["A11:N11"].Merge = true;
            workSheet.Cells["A12:N12"].Merge = true;
            workSheet.Cells["A13:H13"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A5:N5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A6:N6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K8:N8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K10:N10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void TextoNegrita(ExcelWorksheet workSheet)
        {

        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {

        }
    }
}