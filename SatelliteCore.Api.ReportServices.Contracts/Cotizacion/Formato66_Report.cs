using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato66_Report
    {
        public static string Exportar(Image logoUnilene, Formato66_Model cotizacion)
        {
            byte[] file;

            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Almenara");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 5, 0, 5);
                imagenUnilene.SetSize(160, 85);

                worksheet.Cells.Style.Font.Name = "Calibri";
                worksheet.Cells.Style.Font.Size = 11;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                worksheet.Cells["A2"].Value = "ANEXO Nº 05";
                worksheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A2"].Style.Font.Bold = true;
                worksheet.Cells["A2"].Style.Font.Size = 10;

                worksheet.Cells["A3"].Value = "FORMATO DE COTIZACION DE BIENES";
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.Font.Size = 11;

                worksheet.Cells["C4"].Value = "Lima " + cotizacion.Prov_Fecha.ToLongDateString();
                worksheet.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["C4"].Style.Font.Bold = true;
                worksheet.Cells["C4"].Style.Font.Size = 11;


                worksheet.Cells["A5"].Value = "Señores:";
                worksheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A5"].Style.Font.Bold = true;
                worksheet.Cells["A5"].Style.Font.Size = 11;

                worksheet.Cells["I5"].Value = "Cotiz. Nro "+ cotizacion.Prov_NroCotizacion +"-2022 Unilene S.A.C";
                worksheet.Cells["I5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["I5"].Style.Font.Bold = true;
                worksheet.Cells["I5"].Style.Font.Size = 11;

                worksheet.Cells["A6"].Value = "ESSALUD";
                worksheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A6"].Style.Font.Bold = true;
                worksheet.Cells["A6"].Style.Font.Size = 11;

                worksheet.Cells["A7"].Value = "De mi consideracion:";
                worksheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A7"].Style.Font.Size = 11;

                worksheet.Cells["A8"].Value = "En respuesta a la solictud de cotizacion sobre la 'MATERIAL MEDICO DELEDO', y despues de haber analizado las especificaciones tecnicas del mencionado requerimiento," +
                    " las mismas que acepto en todos sus extremos, indico que cumplo con TODOS los requerimientos solicitados.";
                worksheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A8"].Style.Font.Bold = true;
                worksheet.Cells["A8"].Style.Font.Size = 11;
                worksheet.Cells["A8"].Style.WrapText = true;

                worksheet.Cells["A9"].Value = "Asimismo, declaro que las caracteristicas tecnicas de los bienes cotizados por mi representada se ajustan a lo requerido por su Entidad." +
                    " En tal sentido, indico que el costo total por la solucion requerida es la que detallo a continuacion.";
                worksheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A9"].Style.Font.Bold = true;
                worksheet.Cells["A9"].Style.Font.Size = 11;
                worksheet.Cells["A9"].Style.WrapText = true;

                // DETALLE

                worksheet.Cells["A11"].Value = "ITEM";
                worksheet.Cells["A11"].Style.Font.Size = 11;
                worksheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A11"].Style.WrapText = true;

                worksheet.Cells["B11"].Value = "CODIGO SAP";
                worksheet.Cells["B11"].Style.Font.Size = 11;
                worksheet.Cells["B11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B11"].Style.WrapText = true;


                worksheet.Cells["C11"].Value = "DESCRIPCION DEL ITEM";
                worksheet.Cells["C11"].Style.Font.Size = 11;
                worksheet.Cells["C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C11"].Style.WrapText = true;

                worksheet.Cells["D11"].Value = "CANTIDAD TOTAL";
                worksheet.Cells["D11"].Style.Font.Size = 11;
                worksheet.Cells["D11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D11"].Style.WrapText = true;

                worksheet.Cells["E11"].Value = "U.M.";
                worksheet.Cells["E11"].Style.Font.Size = 11;
                worksheet.Cells["E11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E11"].Style.WrapText = true;

                worksheet.Cells["F11"].Value = "PLAZO DE ENTREGA";
                worksheet.Cells["F11"].Style.Font.Size = 11;
                worksheet.Cells["F11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F11"].Style.WrapText = true;

                worksheet.Cells["G11"].Value = "PROCEDENCIA";
                worksheet.Cells["G11"].Style.Font.Size = 11;
                worksheet.Cells["G11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G11"].Style.WrapText = true;


                worksheet.Cells["H11"].Value = "MARCA/MODELO";
                worksheet.Cells["H11"].Style.Font.Size = 11;
                worksheet.Cells["H11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H11"].Style.WrapText = true;

                worksheet.Cells["I11"].Value = "CUMPLE AL 100% CON LAS EE. TT.";
                worksheet.Cells["I11"].Style.Font.Size = 11;
                worksheet.Cells["I11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I11"].Style.WrapText = true;

                worksheet.Cells["J11"].Value = "VIGENCIA DE LA OFERTA";
                worksheet.Cells["J11"].Style.Font.Size = 11;
                worksheet.Cells["J11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J11"].Style.WrapText = true;

                worksheet.Cells["K11"].Value = "RNP";
                worksheet.Cells["K11"].Style.Font.Size = 11;
                worksheet.Cells["K11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K11"].Style.WrapText = true;

                worksheet.Cells["L11"].Value = "PRECIO UNITARIO INC. IGV (S/.)";
                worksheet.Cells["L11"].Style.Font.Size = 11;
                worksheet.Cells["L11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L11"].Style.WrapText = true;

                worksheet.Cells["M11"].Value = "PRECIO TOTAL INC. IGV (S/.)";
                worksheet.Cells["M11"].Style.Font.Size = 11;
                worksheet.Cells["M11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M11"].Style.WrapText = true;


                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                BordesCeldas(worksheet);
                PintarCeldas(worksheet);

                int index = 12;

                foreach (Coti_Formato66_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Row(index).Height = 39.75;
                    worksheet.Cells["A" + index].Value = index - 11;
                    worksheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + index].Style.Numberformat.Format = "0";
                    worksheet.Cells["A" + index].Style.WrapText = true;


                    worksheet.Cells["B" + index].Value = item.CodigoSAP;
                    worksheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + index].Style.WrapText = true;

                    worksheet.Cells["C" + index].Value = item.Descripcion;
                    worksheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + index].Style.WrapText = true;

                    worksheet.Cells["D" + index].Value = item.Cantidad;
                    worksheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + index].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["D" + index].Style.WrapText = true;

                    worksheet.Cells["E" + index].Value = item.Um;
                    worksheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + index].Style.WrapText = true;

                    worksheet.Cells["F" + index].Value = item.PlazoEntrega;
                    worksheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + index].Style.WrapText = true;


                    worksheet.Cells["G" + index].Value = item.Procedencia;
                    worksheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + index].Style.WrapText = true;

                    worksheet.Cells["H" + index].Value = item.Marcamodelo;
                    worksheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + index].Style.WrapText = true;


                    worksheet.Cells["I" + index].Value = item.CumpEETT;
                    worksheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + index].Style.WrapText = true;

                    worksheet.Cells["J" + index].Value = item.VigOferta;
                    worksheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + index].Style.WrapText = true;

                    worksheet.Cells["K" + index].Value = item.Rnp;
                    worksheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + index].Style.WrapText = true;

                    worksheet.Cells["L" + index].Value = item.PreUnitario;
                    worksheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + index].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["L" + index].Style.WrapText = true;

                    worksheet.Cells["M" + index].Value = item.PreTotal;
                    worksheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + index].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["M" + index].Style.WrapText = true;


                    index++;
                }

                worksheet.Row(index).Height = 15;

                worksheet.Cells["A" + index + ":L" + index].Merge = true;
                worksheet.Cells["A" + index + ":L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["A" + index + ":L" + index].Value = "TOTAL S/.";
                worksheet.Cells["A" + index + ":L" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":L" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                worksheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + index].Value = cotizacion.Prov_ValorTotal;
                worksheet.Cells["M" + index].Style.Font.Size = 11;
                worksheet.Cells["M" + index].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["M" + index].Style.Font.Bold = true;
                worksheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["M" + index].Style.WrapText = true;


                index++;
                worksheet.Row(index).Height = 64.5;
                worksheet.Cells["A" + index + ":M" + index].Merge = true;
                worksheet.Cells["A" + index + ":M" + index].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, " +
                    "inpecciones pruebas y, de ser el caso los costos laborales conforme la legislacion vigente, asi como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y/o servicio a" +
                    " contratar, excepto la de aquellos proveedores que gocen de alguna exoneracion legal, no incluiran en el precio de su oferta los tributos respectivos. Asimismo, declaro bajo juramento que, " +
                    "mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, conforme lo establece el articulo 11" +
                    "  del Texto Unico Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado aprobado por Decreto Supremo 082-2019-EF";
                worksheet.Cells["A" + index + ":M" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                index++;
                index++; 
                worksheet.Row(index).Height = 26.25;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "RAZON SOCIAL PROVEEDOR (P. NATURAL O JURIDICA)";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_RazonSocial;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["H" + index + ":I" + index].Merge = true;
                worksheet.Cells["H" + index + ":I" + index].Value = "VIGENCIA - OFERTA" + cotizacion.Prov_vigOferta;
                worksheet.Cells["H" + index + ":I" + index].Style.Font.Size = 11;
                worksheet.Cells["H" + index + ":I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H" + index + ":I" + index].Style.WrapText = true;
                worksheet.Cells["H" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;
                worksheet.Row(index).Height = 26.25;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "RUC";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Ruc;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;
                worksheet.Row(index).Height = 26.25;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "FORMA DE PAGO";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_FormaPago;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                worksheet.Row(index).Height = 26.25;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "CORREO ELECTRONICO";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Email;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                worksheet.Row(index).Height = 26.25;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "TELEFONO FIJO Y MOVIL";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Telefono +" - "+ cotizacion.Prov_movil;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                worksheet.Row(index).Height = 26.25;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "PERSONA DE CONTACTO";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_Contacto;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;

                worksheet.Row(index).Height = 27;
                worksheet.Cells["A" + index + ":C" + index].Merge = true;
                worksheet.Cells["A" + index + ":C" + index].Value = "VIGENCIA DE PRODUCTO";
                worksheet.Cells["A" + index + ":C" + index].Style.Font.Size = 11;
                worksheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["A" + index + ":C" + index].Style.WrapText = true;
                worksheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D" + index + ":G" + index].Merge = true;
                worksheet.Cells["D" + index + ":G" + index].Value = cotizacion.Prov_vigProducto;
                worksheet.Cells["D" + index + ":G" + index].Style.Font.Size = 11;
                worksheet.Cells["D" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + index + ":G" + index].Style.WrapText = true;
                worksheet.Cells["D" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                worksheet.Cells["J" + index + ":M" + index].Merge = true;
                worksheet.Cells["J" + index + ":M" + index].Value = "Firma, nombre y apellidos del proveedor y representante legal o persona autorizada para emitir cotizaciones";
                worksheet.Cells["J" + index + ":M" + index].Style.Font.Size = 11;
                worksheet.Cells["J" + index + ":M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J" + index + ":M" + index].Style.WrapText = true;
               

                worksheet.View.ZoomScale = 80;
                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
              
            }

            return reporte;
        }

   

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 7.00 + 0.71; 
            worksheet.Column(2).Width = 11.57 + 0.71;
            worksheet.Column(3).Width = 45.57 + 0.71;
            worksheet.Column(4).Width = 9.71 + 0.71;
            worksheet.Column(5).Width = 8.43 + 0.71;
            worksheet.Column(6).Width = 14.43 + 0.71;
            worksheet.Column(7).Width = 13.57 + 0.71;
            worksheet.Column(8).Width = 13.29 + 0.71;
            worksheet.Column(9).Width = 9.43  + 0.71;
            worksheet.Column(10).Width = 14.43  + 0.71;
            worksheet.Column(11).Width = 4.57 + 0.71;
            worksheet.Column(12).Width = 13.71  + 0.71;
            worksheet.Column(13).Width = 13.86  + 0.71;

            worksheet.Row(1).Height = 15;
            worksheet.Row(2).Height = 15;
            worksheet.Row(3).Height = 18.75;
            worksheet.Row(4).Height = 15;
            worksheet.Row(5).Height = 24.75;
            worksheet.Row(6).Height = 15;
            worksheet.Row(7).Height = 18.75;

            worksheet.Row(8).Height = 36.75;
            worksheet.Row(9).Height = 41.25;
            worksheet.Row(10).Height = 18.75;
            worksheet.Row(11).Height = 60;



        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A2:M2"].Merge = true;
            worksheet.Cells["A3:M3"].Merge = true;

            worksheet.Cells["A6:B6"].Merge = true;
            worksheet.Cells["A7:B7"].Merge = true;

            

            worksheet.Cells["I5:L5"].Merge = true;

            worksheet.Cells["A8:M8"].Merge = true;
            worksheet.Cells["A9:M9"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A11,B11,C11,D11,E11,F11,G11,H11,I11,J11,K11,L11,M11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["A11:M11"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B4C6E7"));


        }
    }
}
