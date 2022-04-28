using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato71_Report
    {
        public static string Exportar(Image firma,Image logoUnilene, Coti_Formato71_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Sabogal F-2");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 80);

                worksheet.Cells.Style.Font.Size = 10;
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                PintarCeldas(worksheet);
                BordesCeldas(worksheet);
                BordesCeldas(worksheet);
                TextoNegrita(worksheet);

                worksheet.Cells["P1"].Value = "N° COTIZACIÓN Nro. "+ cotizacion.Prov_NroCotizacion;
                worksheet.Cells["P1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["P1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["P1"].Style.Font.Bold = true;
                worksheet.Cells["P1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                worksheet.Cells["A3"].Value = "RUBRO: DISPOSITIVOS MEDICOS";
                worksheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                worksheet.Cells["A4"].Value = "ADQUISICIÓN DE MATERIAL MÉDICO DELEGADO - DICIEMBRE 2021 - PARA EL HOSPITAL ALBERTO SABOGAL SOLOGUREN DE LA RED PRESTACIONAL SABOGAL - GRUPO B";
                worksheet.Cells["A4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A4"].Style.Font.Bold = true;
                worksheet.Cells["A4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));


                worksheet.Cells["A7"].Value = "N°";
                worksheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A7"].Style.Font.Size = 8;
                worksheet.Cells["A7"].Style.WrapText = true;


                worksheet.Cells["C7"].Value = "ESPECIFICACIÓN TÉCNICA DEL PRODUCTO";
                worksheet.Cells["C7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C7"].Style.Font.Bold = true;
                worksheet.Cells["C7"].Style.Font.Size = 8;
                worksheet.Cells["C7"].Style.WrapText = true;

                worksheet.Cells["J7"].Value = "DOCUMENTOS OBLIGATORIOS PARA LA ADMISIÓN DE LA OFERTA (PRESENTACIÓN OBLIGATORIA)";
                worksheet.Cells["J7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J7"].Style.Font.Bold = true;
                worksheet.Cells["J7"].Style.Font.Size = 8;
                worksheet.Cells["J7"].Style.WrapText = true;

                worksheet.Cells["S7"].Value = "REQUISITOS DE CALIFICACIÓN";
                worksheet.Cells["S7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["S7"].Style.Font.Bold = true;
                worksheet.Cells["S7"].Style.Font.Size = 8;
                worksheet.Cells["S7"].Style.WrapText = true;

                worksheet.Cells["U7"].Value = "OTRAS CONDICIONES RELEVANTES";
                worksheet.Cells["U7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["U7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["U7"].Style.Font.Bold = true;
                worksheet.Cells["U7"].Style.Font.Size = 8;
                worksheet.Cells["U7"].Style.WrapText = true;


                worksheet.Cells["B7"].Value = "CÓDIGO SAP";
                worksheet.Cells["B7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B7"].Style.Font.Bold = true;
                worksheet.Cells["B7"].Style.Font.Size = 8;
                worksheet.Cells["B7"].Style.WrapText = true;

                worksheet.Cells["C8"].Value = "DESCRIPCION";
                worksheet.Cells["C8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C8"].Style.Font.Bold = true;
                worksheet.Cells["C8"].Style.Font.Size = 8;
                worksheet.Cells["C8"].Style.WrapText = true;

                worksheet.Cells["D8"].Value = "UNIDADE MEDIDA";
                worksheet.Cells["D8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D8"].Style.Font.Bold = true;
                worksheet.Cells["D8"].Style.Font.Size = 8;
                worksheet.Cells["D8"].Style.WrapText = true;

                worksheet.Cells["E7"].Value = "CANTIDAD";
                worksheet.Cells["E7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E7"].Style.Font.Bold = true;
                worksheet.Cells["E7"].Style.Font.Size = 8;
                worksheet.Cells["E7"].Style.WrapText = true;


                worksheet.Cells["F7"].Value = "MARCA";
                worksheet.Cells["F7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F7"].Style.Font.Bold = true;
                worksheet.Cells["F7"].Style.Font.Size = 8;
                worksheet.Cells["F7"].Style.WrapText = true;

                worksheet.Cells["G7"].Value = "N° DE REGISTO SANITARIO";
                worksheet.Cells["G7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G7"].Style.Font.Bold = true;
                worksheet.Cells["G7"].Style.Font.Size = 8;
                worksheet.Cells["G7"].Style.WrapText = true;


                worksheet.Cells["H7"].Value = "PROCEDENCIA";
                worksheet.Cells["H7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H7"].Style.Font.Bold = true;
                worksheet.Cells["H7"].Style.Font.Size = 8;
                worksheet.Cells["H7"].Style.WrapText = true;

                worksheet.Cells["I7"].Value = "Cumple al 100% con las especificaciones técnicas del producto (Ficha técnica del IETSI)   (SI/NO) ";
                worksheet.Cells["I7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I7"].Style.Font.Bold = true;
                worksheet.Cells["I7"].Style.Font.Size = 8;
                worksheet.Cells["I7"].Style.WrapText = true;

                worksheet.Cells["J8"].Value = "Cuenta con Registro Sanitario o Certificado de Registro Sanitario vigente  (SI/NO) ";
                worksheet.Cells["J8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J8"].Style.Font.Bold = true;
                worksheet.Cells["J8"].Style.Font.Size = 8;
                worksheet.Cells["J8"].Style.WrapText = true;

                worksheet.Cells["K8"].Value = "Cuenta con Certificado de Buenas Prácticas de Manufactura (CBPM) (SI/NO)";
                worksheet.Cells["K8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K8"].Style.Font.Bold = true;
                worksheet.Cells["K8"].Style.Font.Size = 8;
                worksheet.Cells["K8"].Style.WrapText = true;

                worksheet.Cells["L8"].Value = "Certificado de Buenas Prácticas de Almacenamiento (CBPA) (SI/NO)";
                worksheet.Cells["L8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L8"].Style.Font.Bold = true;
                worksheet.Cells["L8"].Style.Font.Size = 8;
                worksheet.Cells["L8"].Style.WrapText = true;

                worksheet.Cells["M8"].Value = "Cuenta con  Certificado de Buenas Prácticas de Distribución y Transporte (SI/NO)";
                worksheet.Cells["M8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M8"].Style.Font.Bold = true;
                worksheet.Cells["M8"].Style.Font.Size = 8;
                worksheet.Cells["M8"].Style.WrapText = true;

                worksheet.Cells["N8"].Value = "Cuenta con  Certificado de Análisis del Dispositivo Médico (SI/NO)";
                worksheet.Cells["N8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N8"].Style.Font.Bold = true;
                worksheet.Cells["N8"].Style.Font.Size = 8;
                worksheet.Cells["N8"].Style.WrapText = true;


                worksheet.Cells["O8"].Value = "Cuenta con   Metodología Analitica (SI/NO)";
                worksheet.Cells["O8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O8"].Style.Font.Bold = true;
                worksheet.Cells["O8"].Style.Font.Size = 8;
                worksheet.Cells["O8"].Style.WrapText = true;

                worksheet.Cells["P8"].Value = "Cuenta con  Ficha Técnica del producto (SI/NO)";
                worksheet.Cells["P8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["P8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["P8"].Style.Font.Bold = true;
                worksheet.Cells["P8"].Style.Font.Size = 8;
                worksheet.Cells["P8"].Style.WrapText = true;

                worksheet.Cells["Q8"].Value = "Presentará muestras en el procedimiento de selección ( SI/NO)";
                worksheet.Cells["Q8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Q8"].Style.Font.Bold = true;
                worksheet.Cells["Q8"].Style.Font.Size = 8;
                worksheet.Cells["Q8"].Style.WrapText = true;


                worksheet.Cells["R8"].Value = "Cuenta con  Manual de Instruccciones de Uso o Inserto (SI/NO)";
                worksheet.Cells["R8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["R8"].Style.Font.Bold = true;
                worksheet.Cells["R8"].Style.Font.Size = 8;
                worksheet.Cells["R8"].Style.WrapText = true;

                worksheet.Cells["S8"].Value = "Cuenta con Resolución de Autorización Sanitaria de Funcionamiento de Establecimiento Farmacéutico (SI/NO)";
                worksheet.Cells["S8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["S8"].Style.Font.Bold = true;
                worksheet.Cells["S8"].Style.Font.Size = 8;
                worksheet.Cells["S8"].Style.WrapText = true;

                worksheet.Cells["T8"].Value = "Cumple con la Experiencia del Postor  señalada en el (Anexo L) (SI/NO)";
                worksheet.Cells["T8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["T8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["T8"].Style.Font.Bold = true;
                worksheet.Cells["T8"].Style.Font.Size = 8;
                worksheet.Cells["T8"].Style.WrapText = true;

                worksheet.Cells["U8"].Value = "Cumple con la Vigencia Mínima de entrega (SI/NO)";
                worksheet.Cells["U8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["U8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["U8"].Style.Font.Bold = true;
                worksheet.Cells["U8"].Style.Font.Size = 8;
                worksheet.Cells["U8"].Style.WrapText = true;

                worksheet.Cells["V8"].Value = "Cumple con los plazos de entrega (SI/NO)";
                worksheet.Cells["V8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["V8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["V8"].Style.Font.Bold = true;
                worksheet.Cells["V8"].Style.Font.Size = 8;
                worksheet.Cells["V8"].Style.WrapText = true;


                worksheet.Cells["W7"].Value = "PRECIO UNITARIO S /.";
                worksheet.Cells["W7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["W7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["W7"].Style.Font.Bold = true;
                worksheet.Cells["W7"].Style.Font.Size = 8;
                worksheet.Cells["W7"].Style.WrapText = true;

                worksheet.Cells["X7"].Value = "MONTO TOTAL INCLUIDO IGV S /.";
                worksheet.Cells["X7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["X7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["X7"].Style.Font.Bold = true;
                worksheet.Cells["X7"].Style.Font.Size = 8;
                worksheet.Cells["X7"].Style.WrapText = true;


                int index = 9;
                foreach (Coti_Formato71_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Row(index).Height = 52.5;
                    worksheet.Cells["A" + index].Value = index - 8;
                    worksheet.Cells["A" + index].Style.Font.Size = 10;
                    worksheet.Cells["A" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + index].Style.Numberformat.Format = "0";
                    worksheet.Cells["A" + index].Style.WrapText = true;


                    worksheet.Cells["B" + index].Value = item.CodigoSAP;
                    worksheet.Cells["B" + index].Style.Font.Size = 10;
                    worksheet.Cells["B" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + index].Style.WrapText = true;


                    worksheet.Cells["C" + index].Value = item.Descripcion;
                    worksheet.Cells["C" + index].Style.Font.Size = 10;
                    worksheet.Cells["C" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + index].Style.WrapText = true;

                    worksheet.Cells["D" + index].Value = item.UndMedida;
                    worksheet.Cells["D" + index].Style.Font.Size = 10;
                    worksheet.Cells["D" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + index].Style.WrapText = true;



                    worksheet.Cells["E" + index].Value = item.Cantidad;
                    worksheet.Cells["E" + index].Style.Font.Size = 10;
                    worksheet.Cells["E" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + index].Style.WrapText = true;
                    worksheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";


                    worksheet.Cells["F" + index].Value = item.Marca;
                    worksheet.Cells["F" + index].Style.Font.Size = 8;
                    worksheet.Cells["F" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + index].Style.WrapText = true;
                    


                    worksheet.Cells["G" + index].Value = item.Rsanitario;
                    worksheet.Cells["G" + index].Style.Font.Size = 8;
                    worksheet.Cells["G" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + index].Style.WrapText = true;


                    worksheet.Cells["H" + index].Value = item.Procedencia;
                    worksheet.Cells["H" + index].Style.Font.Size = 10;
                    worksheet.Cells["H" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + index].Style.WrapText = true;



                    worksheet.Cells["I" + index].Value = item.Fichatecnica;
                    worksheet.Cells["I" + index].Style.Font.Size = 10;
                    worksheet.Cells["I" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + index].Style.WrapText = true;


                    worksheet.Cells["J" + index].Value = item.Crsvigente;
                    worksheet.Cells["J" + index].Style.Font.Size = 10;
                    worksheet.Cells["J" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + index].Style.WrapText = true;


                    worksheet.Cells["K" + index].Value = item.Cbpm;
                    worksheet.Cells["K" + index].Style.Font.Size = 10;
                    worksheet.Cells["K" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + index].Style.WrapText = true;

                    worksheet.Cells["L" + index].Value = item.Cbpa;
                    worksheet.Cells["L" + index].Style.Font.Size = 10;
                    worksheet.Cells["L" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + index].Style.WrapText = true;


                    worksheet.Cells["M" + index].Value = item.Bpdt;
                    worksheet.Cells["M" + index].Style.Font.Size = 10;
                    worksheet.Cells["M" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + index].Style.WrapText = true;


                    worksheet.Cells["N" + index].Value = item.Cmanalisis;
                    worksheet.Cells["N" + index].Style.Font.Size = 10;
                    worksheet.Cells["N" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["N" + index].Style.WrapText = true;


                    worksheet.Cells["O" + index].Value = item.MetodoAnalisis;
                    worksheet.Cells["O" + index].Style.Font.Size = 10;
                    worksheet.Cells["O" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["O" + index].Style.WrapText = true;

                    worksheet.Cells["P" + index].Value = item.Fichatenicaproducto;
                    worksheet.Cells["P" + index].Style.Font.Size = 10;
                    worksheet.Cells["P" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["P" + index].Style.WrapText = true;


                    worksheet.Cells["Q" + index].Value = item.Pmps;
                    worksheet.Cells["Q" + index].Style.Font.Size = 10;
                    worksheet.Cells["Q" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["Q" + index].Style.WrapText = true;


                    worksheet.Cells["R" + index].Value = item.Cuentamanual;
                    worksheet.Cells["R" + index].Style.Font.Size = 10;
                    worksheet.Cells["R" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["R" + index].Style.WrapText = true;


                    worksheet.Cells["S" + index].Value = item.ResAutoSanitaria;
                    worksheet.Cells["S" + index].Style.Font.Size = 10;
                    worksheet.Cells["S" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["S" + index].Style.WrapText = true;


                    worksheet.Cells["T" + index].Value = item.Anexol;
                    worksheet.Cells["T" + index].Style.Font.Size = 10;
                    worksheet.Cells["T" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["T" + index].Style.WrapText = true;



                    worksheet.Cells["U" + index].Value = item.CvigminimaEntrega;
                    worksheet.Cells["U" + index].Style.Font.Size = 10;
                    worksheet.Cells["U" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["U" + index].Style.WrapText = true;

                    worksheet.Cells["V" + index].Value = item.CPlazoEntrega;
                    worksheet.Cells["V" + index].Style.Font.Size = 10;
                    worksheet.Cells["V" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["V" + index].Style.WrapText = true;

                    worksheet.Cells["W" + index].Value = item.PreUnitario;
                    worksheet.Cells["W" + index].Style.Font.Size = 10;
                    worksheet.Cells["W" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["W" + index].Style.WrapText = true;
                    worksheet.Cells["W" + index].Style.Numberformat.Format = "#,##0.00";

                    worksheet.Cells["X" + index].Value = item.PreTotal;
                    worksheet.Cells["X" + index].Style.Font.Size = 10;
                    worksheet.Cells["X" + index].Style.Font.Name = "Calibri";
                    worksheet.Cells["X" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["X" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["X" + index].Style.WrapText = true;
                    worksheet.Cells["X" + index].Style.Numberformat.Format = "#,##0.00";


                    index++;
                }


                worksheet.Cells["W" + index].Value = "TOTAL S/ ";
                worksheet.Cells["W" + index].Style.Font.Size = 10;
                worksheet.Cells["W" + index].Style.Font.Name = "Calibri";
                worksheet.Cells["W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["W" + index].Style.WrapText = true;
               
                worksheet.Cells["X" + index].Value = cotizacion.Prov_valorTotal;
                worksheet.Cells["X" + index].Style.Font.Size = 10;
                worksheet.Cells["X" + index].Style.Font.Name = "Calibri";
                worksheet.Cells["X" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["X" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["X" + index].Style.WrapText = true;
                worksheet.Cells["X" + index].Style.Numberformat.Format = "#,##0.00";


                index = index + 2;
                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "Plazo de entrega : ";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":X" + index].Style.WrapText = true;

                index++;

                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "Nota:";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":X" + index].Style.WrapText = true;


                index++;
                worksheet.Row(index).Height = 29.25;
                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "1.- El precio cotizado es a todo costo, es decir, incluye todos los tributos cuando corresponda (incluido el I.G.V.), seguros, transportes, inspecciones, pruebas y cualquier otro concepto que pueda tener incidencia sobre el costo del bien de acuerdo a las condiciones generales, hasta la entrega en su destino. En caso que el precio deba ser sin IGV (exonerado) indicarlo.";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + " :X" + index].Style.WrapText = true;


                index++;
                worksheet.Row(index).Height = 29.25;
                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "2.- El Valor Estimado Total debe ser expresado como máximo en 2 (DOS) decimales, el precio unitario se podrá expresar como máximo en 3 (TRES) decimales.";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":X" + index].Style.WrapText = true;



                index++;
                worksheet.Row(index).Height = 29.25;
                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "3.- El numeral 42.1 del artículo 42.- Presunción de veracidad de la Ley N° 27444 - Ley del Procedimiento Administrativo General, establece que todas las declaraciones juradas, los documentos sucedáneos presentados y la información incluida en los escritos y formularios que presenten los administrados para la realización de procedimientos administrativos, se presumen verificados por quien hace uso de ellos, así como de contenido veraz para fines del procedimiento administrativo. Esta presunción admite prueba en contrario.";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":X" + index].Style.WrapText = true;


                index++;
                worksheet.Row(index).Height = 36;
                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "4.- Asimismo, el numeral 4 del artículo 56.- Deberes generales de los administrados en el procedimiento del mismo cuerpo legal, estipula como uno de los deberes generales de los administrados, la comprobación de la autenticidad, previamente a su presentación ante la entidad, de la documentación sucedánea y de cualquier otra información que se ampare en la Presunción de Veracidad.";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":X" + index].Style.WrapText = true;

                index++;
                worksheet.Row(index).Height = 21.75;
                worksheet.Cells["A" + index + ":X" + index].Merge = true;
                worksheet.Cells["A" + index + ":X" + index].Value = "5.- El numeral 1.7 Principio de presunción de veracidad.- En la tramitación del procedimiento administrativo, se presume que los documentos y declaraciones formulados por los administrados en la forma prescrita por esta Ley, responden a la verdad de los hechos que ellos afirman. Esta presunción admite prueba en contrario. ";
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":X" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":X" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Razón Social: " + cotizacion.Prov_RazonSocial;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Nº R.U.C.:" + cotizacion.Prov_Ruc;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Representante Legal: " + cotizacion.Prov_Representante;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Contacto:" + cotizacion.Prov_Contacto;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Celular:" + cotizacion.Prov_Celular;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Teléfono: " + cotizacion.Prov_Telefono;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Email:  " + cotizacion.Prov_Email;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Plazo de entrega :  " + cotizacion.Prov_PlazoEntrega;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Cells["A" + index + ":H" + index].Merge = true;
                worksheet.Cells["A" + index + ":H" + index].Value = "Fecha: :  " + cotizacion.Prov_Fecha.ToLongDateString();
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Size = 9;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Bold = true;
                worksheet.Cells["A" + index + ":H" + index].Style.Font.Name = "Arial";
                worksheet.Cells["A" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + index + ":H" + index].Style.WrapText = true;

                index++;
                worksheet.Row(index).Height = 36;
                worksheet.Cells["M" + index + ":P" + index].Merge = true;
                worksheet.Cells["M" + index + ":P" + index].Value = "FIRMA Y SELLO DEL REPRESENTANTE  DE LA EMPRESA";
                worksheet.Cells["M" + index + ":P" + index].Style.Font.Size = 9;
                worksheet.Cells["M" + index + ":P" + index].Style.Font.Bold = true;
                worksheet.Cells["M" + index + ":P" + index].Style.Font.Name = "Arial";
                worksheet.Cells["M" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M" + index + ":P" + index].Style.WrapText = true;

                ExcelPicture firmaCatizacion = worksheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(index-8, -8,11, 85);
                firmaCatizacion.SetSize(265, 120);

                worksheet.View.ZoomScale = 100;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }

            return reporte;
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 2.29 + 0.71;
            worksheet.Column(2).Width = 8.29 + 0.71;
            worksheet.Column(3).Width = 20.43 + 0.71;
            worksheet.Column(4).Width = 6.71 + 0.71;
            worksheet.Column(5).Width = 8.14 + 0.71;
            worksheet.Column(6).Width = 6.29 + 0.71;
            worksheet.Column(7).Width = 8.57 + 0.71;
            worksheet.Column(8).Width = 11.00 + 0.71;
            worksheet.Column(9).Width = 10.00 + 0.71;
            worksheet.Column(10).Width = 10.00 + 0.71;
            worksheet.Column(11).Width = 10.00 + 0.71;
            worksheet.Column(12).Width = 10.00 + 0.71;
            worksheet.Column(13).Width = 10.00 + 0.71;
            worksheet.Column(14).Width = 10.00 + 0.71;
            worksheet.Column(15).Width = 10.00 + 0.71;
            worksheet.Column(16).Width = 11.71 + 0.71;
            worksheet.Column(17).Width = 10.00 + 0.71;
            worksheet.Column(18).Width = 10.00 + 0.71;
            worksheet.Column(19).Width = 10.00 + 0.71;
            worksheet.Column(20).Width = 10.00 + 0.71;
            worksheet.Column(21).Width = 10.00 + 0.71;
            worksheet.Column(22).Width = 10.00 + 0.71;
            worksheet.Column(23).Width = 8.43 + 0.71;
            worksheet.Column(24).Width = 11.00 + 0.71;


            worksheet.Row(1).Height = 9.75;
            worksheet.Row(2).Height = 9.75;
            worksheet.Row(3).Height = 41.25;
            worksheet.Row(4).Height = 32.25;
            worksheet.Row(5).Height = 12.00;
            worksheet.Row(6).Height = 3.00;
            worksheet.Row(7).Height = 27.00;
            worksheet.Row(8).Height = 124.5;
         

        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["P1:R1"].Merge = true;
            worksheet.Cells["A3:X3"].Merge = true;
            worksheet.Cells["A4:X4"].Merge = true;

            worksheet.Cells["A7:A8"].Merge = true;
            worksheet.Cells["B7:B8"].Merge = true;
            worksheet.Cells["C7:D7"].Merge = true;
            worksheet.Cells["J7:R7"].Merge = true;

            worksheet.Cells["S7:T7"].Merge = true;
            worksheet.Cells["U7:V7"].Merge = true;

            worksheet.Cells["E7:E8"].Merge = true;
            worksheet.Cells["F7:F8"].Merge = true;
            worksheet.Cells["G7:G8"].Merge = true;
            worksheet.Cells["H7:H8"].Merge = true;
            worksheet.Cells["I7:I8"].Merge = true;
            worksheet.Cells["W7:W8"].Merge = true;
            worksheet.Cells["X7:X8"].Merge = true;

        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A7:A8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B7:B8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["E7:E8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["F7:F8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["G7:G8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["H7:H8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["I7:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells["C7:D7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["J7:R7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["S7:T7"].Style.Border.BorderAround(ExcelBorderStyle.Thin); 
            worksheet.Cells["U7:V7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["W7:W8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["X7:X8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells["C8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["D8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells["J8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["K8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["L8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["M8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells["O8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["P8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["Q8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["R8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells["S8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["T8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["U8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["V8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {

            worksheet.Cells["A7:X8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
        }
        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
       
        }
    }
}
