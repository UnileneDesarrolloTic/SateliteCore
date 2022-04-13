using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato21_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato21_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Lambayeque");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 50);

                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                workSheet.Cells["A2"].Value = "ANEXO Nº 05";
                workSheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A2"].Style.Font.Bold = true;

                workSheet.Cells["A4"].Value = "FORMATO DE COTIZACION DE BIENES";
                workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A4"].Style.Font.Bold = true;


                workSheet.Cells["A5"].Value = "Cotización N° " + cotizacion.Prov_NroCotizacion + "-" + "2022";
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A5"].Style.Font.Bold = true;
                workSheet.Cells["A5"].Style.Font.Size = 10;
                workSheet.Cells["A5"].Style.Font.UnderLine = true;

                workSheet.Cells["A7"].Value = "Lima, "+cotizacion.Fecha_Cotizacion.ToLongDateString() +"\n"+
                "Señores ESSALUD LAMBAYEQUE \n" +
                "De mi consideración: \n" +
                "En respuesta a la solicitud de cotización sobre la la Compra de MATERIAL MEDICO, y después de haber analizado las especificaciones técnicas del mencionado requerimiento, las mismas que acepto en todos sus extremos, indico que cumplo con TODOS los requerimientos solicitados. " +
                "Asimismo, declaro que las características técnicas de los bienes cotizados por mi representada sLim ntido, indico que el costo total por la solución requerida es la que detallo a continuación: ";
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A7"].Style.Font.Bold = true;
                workSheet.Cells["A7"].Style.Font.Size = 10;
                workSheet.Cells["A7"].Style.WrapText = true;

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);


                workSheet.Cells["A9"].Value = "DESCRIPCIÓN DE LOS ÍTEMS";
                workSheet.Cells["A9"].Style.Font.Size = 10;
                workSheet.Cells["A9"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A9"].Style.Font.Bold = true;
                workSheet.Cells["A9"].Style.WrapText = true;

                workSheet.Cells["L9"].Value = "REQUERIMIENTOS MÍNIMOS EN BASES";
                workSheet.Cells["L9"].Style.Font.Size = 10;
                workSheet.Cells["L9"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["L9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L9"].Style.Font.Bold = true;
                workSheet.Cells["L9"].Style.WrapText = true;


                workSheet.Cells["A10"].Value = "N° ITEM";
                workSheet.Cells["A10"].Style.Font.Size = 10;
                workSheet.Cells["A10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A10"].Style.TextRotation = 90;
                workSheet.Cells["A10"].Style.Font.Bold =true;
                workSheet.Cells["A10"].Style.WrapText = true;


                workSheet.Cells["B10"].Value = "CÓDIGO SAP";
                workSheet.Cells["B10"].Style.Font.Size = 10;
                workSheet.Cells["B10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B10"].Style.Font.Bold = true;
                workSheet.Cells["B10"].Style.WrapText = true;


                workSheet.Cells["C10"].Value = "DESCRIPCIÓN DE LOS ÍTEMS";
                workSheet.Cells["C10"].Style.Font.Size = 10;
                workSheet.Cells["C10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C10"].Style.Font.Bold = true;
                workSheet.Cells["C10"].Style.WrapText = true;

                workSheet.Cells["D10"].Value = "Und Med";
                workSheet.Cells["D10"].Style.Font.Size = 10;
                workSheet.Cells["D10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D10"].Style.Font.Bold = true;
                workSheet.Cells["D10"].Style.WrapText = true;
                workSheet.Cells["D10"].Style.TextRotation = 90;

                workSheet.Cells["E10"].Value = "Cant.";
                workSheet.Cells["E10"].Style.Font.Size = 10;
                workSheet.Cells["E10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E10"].Style.Font.Bold = true;
                workSheet.Cells["E10"].Style.WrapText = true;
                workSheet.Cells["E10"].Style.TextRotation = 90;

                workSheet.Cells["F10"].Value = "PRESENTACION DEL PRODUCTO";
                workSheet.Cells["F10"].Style.Font.Size = 10;
                workSheet.Cells["F10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F10"].Style.Font.Bold = true;
                workSheet.Cells["F10"].Style.WrapText = true;
                workSheet.Cells["F10"].Style.TextRotation = 90;

                workSheet.Cells["G10"].Value = "PRECIO UNITARIO (SOLES) INCLUIDO IGV";
                workSheet.Cells["G10"].Style.Font.Size = 10;
                workSheet.Cells["G10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G10"].Style.Font.Bold = true;
                workSheet.Cells["G10"].Style.WrapText = true;

                workSheet.Cells["H10"].Value = "PRECIO TOTAL (SOLES) INCLUIDO IGV";
                workSheet.Cells["H10"].Style.Font.Size = 10;
                workSheet.Cells["H10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H10"].Style.Font.Bold = true;
                workSheet.Cells["H10"].Style.WrapText = true;

                workSheet.Cells["I10"].Value = "MARCA";
                workSheet.Cells["I10"].Style.Font.Size = 10;
                workSheet.Cells["I10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["I10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I10"].Style.Font.Bold = true;
                workSheet.Cells["I10"].Style.WrapText = true;


                workSheet.Cells["J10"].Value = "PAIS DE PROCEDENCIA";
                workSheet.Cells["J10"].Style.Font.Size = 10;
                workSheet.Cells["J10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["J10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J10"].Style.Font.Bold = true;
                workSheet.Cells["J10"].Style.WrapText = true;

                workSheet.Cells["K10"].Value = "VALIDEZ DE LA OFERTA Ò VIGENCIA DE LA COTIZACION";
                workSheet.Cells["K10"].Style.Font.Size = 10;
                workSheet.Cells["K10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["K10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K10"].Style.Font.Bold = true;
                workSheet.Cells["K10"].Style.WrapText = true;
                workSheet.Cells["K10"].Style.TextRotation = 90;

                workSheet.Cells["L10"].Value = "CALIDAD";
                workSheet.Cells["L10"].Style.Font.Size = 10;
                workSheet.Cells["L10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L10"].Style.Font.Bold = true;
                workSheet.Cells["L10"].Style.WrapText = true;


                workSheet.Cells["V10"].Value = "OPORTUNIDAD";
                workSheet.Cells["V10"].Style.Font.Size = 10;
                workSheet.Cells["V10"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["V10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["V10"].Style.Font.Bold = true;
                workSheet.Cells["V10"].Style.WrapText = true;

                workSheet.Cells["L11"].Value = "La vigencia mínima del Material Mèdico deberá ser de 18 meses (indicar vigencia,  en caso de no cumplir con lo solicitado debera adjuntar Carta compromiso de Canje.";
                workSheet.Cells["L11"].Style.Font.Size = 10;
                workSheet.Cells["L11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["L11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L11"].Style.Font.Bold = true;
                workSheet.Cells["L11"].Style.WrapText = true;
                workSheet.Cells["L11"].Style.TextRotation = 90;

                workSheet.Cells["M11"].Value = "Cumple al 100% con la denominación y descripción del ítem ";
                workSheet.Cells["M11"].Style.Font.Size = 10;
                workSheet.Cells["M11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["M11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M11"].Style.Font.Bold = true;
                workSheet.Cells["M11"].Style.WrapText = true;
                workSheet.Cells["M11"].Style.TextRotation = 90;

                workSheet.Cells["N11"].Value = "Cuenta con Resolución de Autorización Sanitaria de Funcionamiento de Establecimiento Farmacéutico- vigente(Si o No)";
                workSheet.Cells["N11"].Style.Font.Size = 10;
                workSheet.Cells["N11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["N11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N11"].Style.Font.Bold = true;
                workSheet.Cells["N11"].Style.WrapText = true;
                workSheet.Cells["N11"].Style.TextRotation = 90;

                workSheet.Cells["O11"].Value = "Certificado de Buena Práctica de Almacenamiento – (CBPA) vigente (Si o No)";
                workSheet.Cells["O11"].Style.Font.Size = 10;
                workSheet.Cells["O11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["O11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O11"].Style.Font.Bold = true;
                workSheet.Cells["O11"].Style.WrapText = true;
                workSheet.Cells["O11"].Style.TextRotation = 90;

                workSheet.Cells["P11"].Value = "Cuenta con Registro Sanitario o Certificado de Registro Sanitario vigente (SI o NO)";
                workSheet.Cells["P11"].Style.Font.Size = 10;
                workSheet.Cells["P11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["P11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P11"].Style.Font.Bold = true;
                workSheet.Cells["P11"].Style.WrapText = true;
                workSheet.Cells["P11"].Style.TextRotation = 90;

                workSheet.Cells["Q11"].Value = "Nº de Registro Sanitario";
                workSheet.Cells["Q11"].Style.Font.Size = 12;
                workSheet.Cells["Q11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["Q11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q11"].Style.Font.Bold = true;
                workSheet.Cells["Q11"].Style.WrapText = true;
                workSheet.Cells["Q11"].Style.TextRotation = 90;

                workSheet.Cells["R11"].Value = "Certificado de Buena Práctica de Manufactura del laboratorio fabricante – (CBPM) vigente.(la fecha de emisión no deberá ser mayor a dos (02) años) ó (*)Alternativos(SI  ó  NO)";
                workSheet.Cells["R11"].Style.Font.Size = 12;
                workSheet.Cells["R11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["R11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R11"].Style.Font.Bold = true;
                workSheet.Cells["R11"].Style.WrapText = true;
                workSheet.Cells["R11"].Style.TextRotation = 90;

                workSheet.Cells["S11"].Value = "Cuenta con Certificado de Análisis del producto terminado (Protocolo de Análisis) (SI o NO)";
                workSheet.Cells["S11"].Style.Font.Size = 12;
                workSheet.Cells["S11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["S11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S11"].Style.Font.Bold = true;
                workSheet.Cells["S11"].Style.WrapText = true;
                workSheet.Cells["S11"].Style.TextRotation = 90;

                workSheet.Cells["T11"].Value = "Cuenta con Metodolodia de Analisis (SI o NO)";
                workSheet.Cells["T11"].Style.Font.Size = 12;
                workSheet.Cells["T11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["T11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T11"].Style.Font.Bold = true;
                workSheet.Cells["T11"].Style.WrapText = true;
                workSheet.Cells["T11"].Style.TextRotation = 90;

                workSheet.Cells["U11"].Value = "Envase mediato e inmediato lleva el logotipo solicitado por la Entidad (SI o NO)";
                workSheet.Cells["U11"].Style.Font.Size = 12;
                workSheet.Cells["U11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["U11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["U11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["U11"].Style.Font.Bold = true;
                workSheet.Cells["U11"].Style.WrapText = true;
                workSheet.Cells["U11"].Style.TextRotation = 90;


                workSheet.Cells["V11"].Value = "Cumple con los Plazos de Entrega establecidos.";
                workSheet.Cells["V11"].Style.Font.Size = 12;
                workSheet.Cells["V11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["V11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["V11"].Style.Font.Bold = true;
                workSheet.Cells["V11"].Style.WrapText = true;
                workSheet.Cells["V11"].Style.TextRotation = 90;

                workSheet.Cells["W11"].Value = "Capacidad de Atención al 100% de lo solicitado";
                workSheet.Cells["W11"].Style.Font.Size = 12;
                workSheet.Cells["W11"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["W11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["W11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["W11"].Style.Font.Bold = true;
                workSheet.Cells["W11"].Style.WrapText = true;
                workSheet.Cells["W11"].Style.TextRotation = 90;

                int index = 13;

                foreach (Coti_Formato21_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Row(index).Height = 63.75;
                    workSheet.Cells["A" + index].Value = index - 12;
                    workSheet.Cells["A" + index].Style.Font.Size = 9;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.WrapText = true;
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                   
                    workSheet.Cells["B" + index].Value = item.CodigoSap;
                    workSheet.Cells["B" + index].Style.Font.Size = 9;
                    workSheet.Cells["B" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;

                   
                    workSheet.Cells["C" + index].Value = item.Descripcion;
                    workSheet.Cells["C" + index].Style.Font.Size = 9;
                    workSheet.Cells["C" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;

                    
                    workSheet.Cells["D" + index].Value = item.UndMedida;
                    workSheet.Cells["D" + index].Style.Font.Size = 9;
                    workSheet.Cells["D" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;

                   
                    workSheet.Cells["E" + index].Value = item.Cantidad;
                    workSheet.Cells["E" + index].Style.Font.Size = 9;
                    workSheet.Cells["E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;
                    workSheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";

                    
                    workSheet.Cells["F" + index].Value = item.Presentacion;
                    workSheet.Cells["F" + index].Style.Font.Size = 9;
                    workSheet.Cells["F" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;


                    
                    workSheet.Cells["G" + index].Value = item.PreUnitario;
                    workSheet.Cells["G" + index].Style.Font.Size = 9;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;
                    workSheet.Cells["G" + index].Style.Numberformat.Format = "#,##0.00";

                    
                    workSheet.Cells["H" + index].Value = item.PreTotal;
                    workSheet.Cells["H" + index].Style.Font.Size = 9;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;
                    workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0.00";

                    
                    workSheet.Cells["I" + index].Value = item.Marca;
                    workSheet.Cells["I" + index].Style.Font.Size = 9;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Row(index).Height = 63.75;
                    workSheet.Cells["J" + index].Value = item.Procedencia;
                    workSheet.Cells["J" + index].Style.Font.Size = 9;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;

                    
                    workSheet.Cells["K" + index].Value = item.ValidezOferta;
                    workSheet.Cells["K" + index].Style.Font.Size = 9;
                    workSheet.Cells["K" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;

                    workSheet.Row(index).Height = 63.75;
                    workSheet.Cells["L" + index].Value = item.VigMinima;
                    workSheet.Cells["L" + index].Style.Font.Size = 9;
                    workSheet.Cells["L" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;

                    
                    workSheet.Cells["M" + index].Value = item.CumpDescripcion;
                    workSheet.Cells["M" + index].Style.Font.Size = 9;
                    workSheet.Cells["M" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    
                    workSheet.Cells["N" + index].Value = item.AutSanitaria;
                    workSheet.Cells["N" + index].Style.Font.Size = 9;
                    workSheet.Cells["N" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;

                   
                    workSheet.Cells["O" + index].Value = item.BuenasPracticas;
                    workSheet.Cells["O" + index].Style.Font.Size = 9;
                    workSheet.Cells["O" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;

                   
                    workSheet.Cells["P" + index].Value = item.CertSanitario;
                    workSheet.Cells["P" + index].Style.Font.Size = 9;
                    workSheet.Cells["P" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["P" + index].Style.WrapText = true;


                   
                    workSheet.Cells["Q" + index].Value = item.NrRegSanitario;
                    workSheet.Cells["Q" + index].Style.Font.Size = 9;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;
                    workSheet.Cells["Q" + index].Style.TextRotation = 90;

                    workSheet.Cells["R" + index].Value = item.CertManufactura;
                    workSheet.Cells["R" + index].Style.Font.Size = 9;
                    workSheet.Cells["R" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["R" + index].Style.WrapText = true;

                    workSheet.Cells["S" + index].Value = item.CertProductoTer;
                    workSheet.Cells["S" + index].Style.Font.Size = 9;
                    workSheet.Cells["S" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["S" + index].Style.WrapText = true;


                    workSheet.Cells["T" + index].Value = item.MetodoAnalisis;
                    workSheet.Cells["T" + index].Style.Font.Size = 9;
                    workSheet.Cells["T" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["T" + index].Style.WrapText = true;


                    workSheet.Cells["U" + index].Value = item.LogSolicitado;
                    workSheet.Cells["U" + index].Style.Font.Size = 9;
                    workSheet.Cells["U" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["U" + index].Style.WrapText = true;

                    workSheet.Cells["V" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["V" + index].Style.Font.Size = 9;
                    workSheet.Cells["V" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["V" + index].Style.WrapText = true;

                    workSheet.Cells["W" + index].Value = item.CapacidadAtencion;
                    workSheet.Cells["W" + index].Style.Font.Size = 9;
                    workSheet.Cells["W" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["W" + index].Style.WrapText = true;

                    index++;

                }

                workSheet.Cells["G" + index].Value = "TOTAL S/";
                workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["H" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0.00";

                index++;
                workSheet.Row(index).Height = 15;
                workSheet.Cells["A" + index + ":W" + index].Merge = true;
                workSheet.Cells["A" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "VALOR TOTAL DE LA COTIZACIÓN";
                workSheet.Cells["A" + index + ":W" + index].Style.Font.Size = 11;
                workSheet.Cells["A" + index + ":W" + index].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A" + index + ":W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index + ":W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":W" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 58.5;
                workSheet.Cells["A" + index + ":W" + index].Merge = true;
                workSheet.Cells["A" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["A" + index].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso, los costos laborales conforme la legislación vigente, así como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y/o servicio a contratar; excepto la de aquellos proveedores que gocen de alguna exoneración legal, no incluirán en el precio de su oferta los tributos respectivos. Asimismo, declaro bajo juramente que, mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, conforme lo establece el artículo 11 del Texto Único Ordenado de la Ley N° 30225, Ley de Contrataciones del Estado, aprobado por Decreto Supremo N° 082-2019-EF.";
                workSheet.Cells["A" + index + ":W" + index].Style.Font.Size = 11;
                workSheet.Cells["A" + index + ":W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index + ":W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["A" + index + ":W" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index].Value = "RAZÓN SOCIAL";
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "Nº RUC";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "PLAZO DE ENTREGA";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "FORMA DE PAGO";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_FormaPago;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "GARANTÍA";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_Garantia + " Vigencia del Producto : " + cotizacion.Prov_VigProducto;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "CORREO ELECTRÓNICO";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_Email;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "TELÉFONO FIJO ";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "PERSONA DE CONTACTO ";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "TELÉFONO MÓVIL";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_Celular;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                index++;
                workSheet.Cells["A" + index + ":C" + index].Merge = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "VIGENCIA DE OFERTA";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F" + index + ":W" + index].Merge = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Font.Bold = true;
                workSheet.Cells["F" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F" + index].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

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
            workSheet.Column(1).Width = 3.43 + 0.71;
            workSheet.Column(2).Width = 11.29 + 0.71;
            workSheet.Column(3).Width = 24.00 + 0.71;
            workSheet.Column(4).Width = 2.86 + 0.71;
            workSheet.Column(5).Width = 6.00 + 0.71;
            workSheet.Column(6).Width = 7.43 + 0.71;
            workSheet.Column(7).Width = 9.57 + 0.71;
            workSheet.Column(8).Width = 12.71 + 0.71;
            workSheet.Column(9).Width = 7.86 + 0.71;
            workSheet.Column(10).Width = 8.57 + 0.71;
            workSheet.Column(11).Width = 9.43 + 0.71;
            workSheet.Column(12).Width = 13.57 + 0.71;
            workSheet.Column(13).Width = 6.57 + 0.71;
            workSheet.Column(14).Width = 6.57 + 0.71;
            workSheet.Column(15).Width = 8.29 + 0.71;
            workSheet.Column(16).Width = 8.14 + 0.71;
            workSheet.Column(17).Width = 5.14 + 0.71;
            workSheet.Column(18).Width = 8.57 + 0.71;
            workSheet.Column(19).Width = 7.57 + 0.71;
            workSheet.Column(20).Width = 10.29 + 0.71;
            workSheet.Column(21).Width = 10.29 + 0.71;
            workSheet.Column(22).Width = 9.86 + 0.71;
            workSheet.Column(22).Width = 7.86 + 0.71;


            workSheet.Row(1).Height = 15.00;
            workSheet.Row(2).Height = 15.00;
            workSheet.Row(3).Height = 8.25;
            workSheet.Row(4).Height = 24.00;
            workSheet.Row(5).Height = 18.75;
            workSheet.Row(6).Height = 9;
            workSheet.Row(7).Height = 88.5;
            workSheet.Row(8).Height = 6;
            workSheet.Row(9).Height = 18;

            workSheet.Row(10).Height = 17.25;
            workSheet.Row(11).Height = 25.5;
            workSheet.Row(12).Height = 118.5;



        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A2:W2"].Merge = true;
            workSheet.Cells["A4:W4"].Merge = true;
            workSheet.Cells["A5:C5"].Merge = true;
            workSheet.Cells["A7:W7"].Merge = true;
            workSheet.Cells["A9:K9"].Merge = true;
            workSheet.Cells["L9:W9"].Merge = true;

            workSheet.Cells["A10:A12"].Merge = true;
            workSheet.Cells["B10:B12"].Merge = true;
            workSheet.Cells["C10:C12"].Merge = true;
            workSheet.Cells["D10:D12"].Merge = true;
            workSheet.Cells["E10:E12"].Merge = true;
            workSheet.Cells["F10:F12"].Merge = true;
            workSheet.Cells["G10:G12"].Merge = true;
            workSheet.Cells["H10:H12"].Merge = true;
            workSheet.Cells["I10:I12"].Merge = true;
            workSheet.Cells["J10:J12"].Merge = true;
            workSheet.Cells["K10:K12"].Merge = true;
            workSheet.Cells["L10:U10"].Merge = true;
            workSheet.Cells["V10:W10"].Merge = true;

            workSheet.Cells["L11:L12"].Merge = true;
            workSheet.Cells["M11:M12"].Merge = true;
            workSheet.Cells["N11:N12"].Merge = true;
            workSheet.Cells["O11:O12"].Merge = true;
            workSheet.Cells["P11:P12"].Merge = true;

            workSheet.Cells["Q11:Q12"].Merge = true;
            workSheet.Cells["R11:R12"].Merge = true;
            workSheet.Cells["S11:S12"].Merge = true;

            workSheet.Cells["T11:T12"].Merge = true;

            workSheet.Cells["U11:U12"].Merge = true;
            workSheet.Cells["V11:V12"].Merge = true;
            workSheet.Cells["W11:W12"].Merge = true;
        }
        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A4:W4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A9:K9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L9:W9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A10:A12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B10:B12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C10:C12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D10:D12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E10:E12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F10:F12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G10:G12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H10:H12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I10:I12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J10:J12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K10:K12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L10:U10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["V10:W10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L11:L12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M11:M12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N11:N12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O11:O12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P11:P12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q11:Q12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R11:R12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S11:S12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["T11:T12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["U11:U12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["V11:V12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["W11:W12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
           
            workSheet.Cells["A9:K9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["L9:W9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["A10:A12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["B10:B12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["C10:C12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["D10:D12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["E10:E12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["F10:F12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["G10:G12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["H10:H12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["I10:I12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["J10:J12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["K10:K12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["L10:U10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["V10:W10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));

            workSheet.Cells["L11:L12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["M11:M12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["N11:N12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["O11:O12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["P11:P12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["Q11:Q12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["R11:R12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["S11:S12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["T11:T12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["U11:U12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["V11:V12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));
            workSheet.Cells["W11:W12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A3D1FB"));

        }
        private static void TextoNegrita(ExcelWorksheet workSheet)
        {

        }

    }
}
