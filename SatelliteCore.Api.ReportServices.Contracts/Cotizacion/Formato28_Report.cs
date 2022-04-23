using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato28_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Formato28_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Almenara");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(2, 20, 0, 7);
                imagenUnilene.SetSize(143, 78);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 10;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);

                worksheet.Cells["A1"].Value = "  RED ASISTENCIAL PUNO - ESSALUD";
                worksheet.Cells["A1"].Style.Font.Size = 8;
                worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A2"].Value = "  DIVISIÓN DE ADQUISICIONES";
                worksheet.Cells["A2"].Style.Font.Size = 8;
                worksheet.Cells["A2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["Q2"].Value = "N°";
                worksheet.Cells["Q2"].Style.Font.Size = 14;
                worksheet.Cells["Q2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["R2"].Value = cotizacion.Nro_Cotizacion;
                worksheet.Cells["R2"].Style.Font.Size = 14;
                worksheet.Cells["R2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F5"].Value = "DECLARACION JURADA DE SOLICITUD DE COTIZACION ADQUISICION DE MATERIAL MEDICO DELEGADO - A COMPRA LOCAL 2022 PARA LA RED ASISTENCIAL PUNO";
                worksheet.Cells["F5"].Style.Font.Size = 13;
                worksheet.Cells["F5"].Style.WrapText = true;
                worksheet.Cells["F5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["Q5"].Value = "FECHA";
                worksheet.Cells["Q5"].Style.Font.Size = 10;
                worksheet.Cells["Q5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["Q6"].Value = cotizacion.Fecha_documento;
                worksheet.Cells["Q6"].Style.Font.Size = 13;
                worksheet.Cells["Q6"].Style.Numberformat.Format = "dd";
                worksheet.Cells["Q6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["R6"].Value = cotizacion.Fecha_documento;
                worksheet.Cells["R6"].Style.Font.Size = 13;
                worksheet.Cells["R6"].Style.Numberformat.Format = "MM";
                worksheet.Cells["R6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["S6"].Value = cotizacion.Fecha_documento;
                worksheet.Cells["S6"].Style.Font.Size = 13;
                worksheet.Cells["S6"].Style.Numberformat.Format = "yyyy";
                worksheet.Cells["S6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A8"].Value = "RAZÓN SOCIAL";
                worksheet.Cells["A8"].Style.Font.Size = 8;
                worksheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E8"].Value = cotizacion.Razon_social;
                worksheet.Cells["E8"].Style.Font.Size = 8;
                worksheet.Cells["E8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A9"].Value = "R.U.C.";
                worksheet.Cells["A9"].Style.Font.Size = 8;
                worksheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E9"].Value = cotizacion.Ruc;
                worksheet.Cells["E9"].Style.Font.Size = 8;
                worksheet.Cells["E9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A10"].Value = "DIRECCIÓN";
                worksheet.Cells["A10"].Style.Font.Size = 8;
                worksheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E10"].Value = cotizacion.Direccion;
                worksheet.Cells["E10"].Style.Font.Size = 8;
                worksheet.Cells["E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A11"].Value = "TELÉFONO";
                worksheet.Cells["A11"].Style.Font.Size = 8;
                worksheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E11"].Value = cotizacion.Telefono;
                worksheet.Cells["E11"].Style.Font.Size = 8;
                worksheet.Cells["E11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A12"].Value = "E-MAIL";
                worksheet.Cells["A12"].Style.Font.Size = 8;
                worksheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E12"].Value = cotizacion.Email_unilene;
                worksheet.Cells["E12"].Style.Font.Size = 8;
                worksheet.Cells["E12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A13"].Value = "VIGENCIA DE OFERTA";
                worksheet.Cells["A13"].Style.Font.Size = 8;
                worksheet.Cells["A13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E13"].Value = cotizacion.Vigencia;
                worksheet.Cells["E13"].Style.Font.Size = 8;
                worksheet.Cells["E13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["L8"].Value = "REPRESENTANTE";
                worksheet.Cells["L8"].Style.Font.Size = 8;
                worksheet.Cells["L8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N8"].Value = cotizacion.Representante;
                worksheet.Cells["N8"].Style.Font.Size = 8;
                worksheet.Cells["N8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["L9"].Value = "CARGO";
                worksheet.Cells["L9"].Style.Font.Size = 8;
                worksheet.Cells["L9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N9"].Value = cotizacion.Cargo;
                worksheet.Cells["N9"].Style.Font.Size = 8;
                worksheet.Cells["N9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["L10"].Value = "E-MAIL";
                worksheet.Cells["L10"].Style.Font.Size = 8;
                worksheet.Cells["L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N10"].Value = cotizacion.Email_representante;
                worksheet.Cells["N10"].Style.Font.Size = 8;
                worksheet.Cells["N10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["L11"].Value = "CELULAR";
                worksheet.Cells["L11"].Style.Font.Size = 8;
                worksheet.Cells["L11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N11"].Value = cotizacion.Celular;
                worksheet.Cells["N11"].Style.Font.Size = 8;
                worksheet.Cells["N11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["L12"].Value = "RPM / RPC";
                worksheet.Cells["L12"].Style.Font.Size = 8;
                worksheet.Cells["L12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N12"].Value = cotizacion.Rpm_rpc;
                worksheet.Cells["N12"].Style.Font.Size = 8;
                worksheet.Cells["N12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["L13"].Value = "CONDICIÓN DE PAGO";
                worksheet.Cells["L13"].Style.Font.Size = 8;
                worksheet.Cells["L13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N13"].Value = cotizacion.Codicion_pago;
                worksheet.Cells["N13"].Style.Font.Size = 8;
                worksheet.Cells["N13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A15"].Value = "SIRVASE LLENAR LA SOLICITUD DE COTIZACION PARA COMPRA DE:";
                worksheet.Cells["A15"].Style.Font.Size = 8;
                worksheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A17"].Value = "N° ITEM";
                worksheet.Cells["A17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A17"].Style.WrapText = true;

                worksheet.Cells["B24"].Value = "SOL PEDIDO";
                worksheet.Cells["B24"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B24"].Style.WrapText = true;

                worksheet.Cells["C24"].Value = "POS";
                worksheet.Cells["C24"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C24"].Style.WrapText = true;

                worksheet.Cells["D17"].Value = "DESCRIPCIÓN DEL ÍTEM";
                worksheet.Cells["D17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D17"].Style.WrapText = true;

                worksheet.Cells["D22"].Value = "CÓDIGO SAP";
                worksheet.Cells["D22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D22"].Style.WrapText = true;

                worksheet.Cells["E22"].Value = "DENOMINACIÓN";
                worksheet.Cells["E22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F22"].Value = "UM";
                worksheet.Cells["F22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G22"].Value = "CANTIDAD";
                worksheet.Cells["G22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G22"].Style.WrapText = true;

                worksheet.Cells["H17"].Value = "REQUISITOS TÉCNICOS MÍNIMOS Y CONDICIONES GEN. P/ADQUISICIÓN DE DISPOSITIVO MEDICO EN ESSALUD";
                worksheet.Cells["H17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H17"].Style.WrapText = true;

                worksheet.Cells["H18"].Value = "MARCA DEL PRODUCTO";
                worksheet.Cells["H18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H18"].Style.TextRotation = 90;
                worksheet.Cells["H18"].Style.WrapText = true;

                worksheet.Cells["I18"].Value = "PROCEDENCIA";
                worksheet.Cells["I18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I18"].Style.TextRotation = 90;
                worksheet.Cells["I18"].Style.WrapText = true;

                worksheet.Cells["J18"].Value = "Forma de Presentación del Producto";
                worksheet.Cells["J18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J18"].Style.TextRotation = 90;
                worksheet.Cells["J18"].Style.WrapText = true;

                worksheet.Cells["K18"].Value = "DOCUMENTOS TÉCNICOS A ADJUNTAR - copia simple  (obligatorio)";
                worksheet.Cells["K18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K18"].Style.WrapText = true;

                worksheet.Cells["K19"].Value = "DEL POSTOR";
                worksheet.Cells["K19"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K19"].Style.WrapText = true;

                worksheet.Cells["M19"].Value = "DEL DISPOSITIVO MEDICO";
                worksheet.Cells["M19"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M19"].Style.WrapText = true;

                worksheet.Cells["K20"].Value = "Resolución de Autorización Sanitaria de Funcionamiento de Establecimiento Farmac. o const. de Reg. de Estab. Farm. (Si o No)";
                worksheet.Cells["K20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K20"].Style.TextRotation = 90;
                worksheet.Cells["K20"].Style.WrapText = true;

                worksheet.Cells["L20"].Value = "Certificado de Buena Práctica de Almacenamiento – (BPA) vigente (Si o No)";
                worksheet.Cells["L20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L20"].Style.TextRotation = 90;
                worksheet.Cells["L20"].Style.WrapText = true;

                worksheet.Cells["M20"].Value = "Registro Sanitario o Certificado de Registro Sanitario (Si o No)";
                worksheet.Cells["M20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M20"].Style.TextRotation = 90;
                worksheet.Cells["M20"].Style.WrapText = true;

                worksheet.Cells["N20"].Value = "Certificado de Buenas Prácticas de Manufactura (Si o No)";
                worksheet.Cells["N20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N20"].Style.TextRotation = 90;
                worksheet.Cells["N20"].Style.WrapText = true;

                worksheet.Cells["O20"].Value = "Metodología de análisis (Si o No)";
                worksheet.Cells["O20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O20"].Style.TextRotation = 90;
                worksheet.Cells["O20"].Style.WrapText = true;

                worksheet.Cells["P20"].Value = "Envase y condiciones de almacenamiento - LOGOTIPO";
                worksheet.Cells["P20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["P20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["P20"].Style.TextRotation = 90;
                worksheet.Cells["P20"].Style.WrapText = true;

                worksheet.Cells["Q17"].Value = "PLAZO DE ENTREGA EN DIAS CALENADARIO";
                worksheet.Cells["Q17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Q17"].Style.TextRotation = 90;
                worksheet.Cells["Q17"].Style.WrapText = true;

                worksheet.Cells["R17"].Value = "PRECIO UNITARIO A OFERTAR\nS /.";
                worksheet.Cells["R17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["R17"].Style.WrapText = true;

                worksheet.Cells["S17"].Value = "VALOR TOTAL A OFERTAR\nS /.";
                worksheet.Cells["S17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["S17"].Style.WrapText = true;

                int row = 25;

                foreach (Formato28_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Cells["A" + row].Value = row - 24;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["B" + row].Value = item.Sol_pedido;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.WrapText = true;

                    worksheet.Cells["C" + row].Value = item.Pos;
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.WrapText = true;

                    worksheet.Cells["D" + row].Value = item.Codigo_sap;
                    worksheet.Cells["D" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.WrapText = true;

                    worksheet.Cells["E" + row].Value = item.Denominacion;
                    worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.WrapText = true;

                    worksheet.Cells["F" + row].Value = item.Um;
                    worksheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.WrapText = true;

                    worksheet.Cells["G" + row].Value = item.Cantidad;
                    worksheet.Cells["G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.WrapText = true;

                    worksheet.Cells["H" + row].Value = item.Marca;
                    worksheet.Cells["H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.WrapText = true;

                    worksheet.Cells["I" + row].Value = item.Procedencia;
                    worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.WrapText = true;

                    worksheet.Cells["J" + row].Value = item.Presentacion;
                    worksheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.WrapText = true;

                    worksheet.Cells["K" + row].Value = item.Autorizacion_funcionamiento;
                    worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["k" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.WrapText = true;

                    worksheet.Cells["L" + row].Value = item.Bp_almacenamiento;
                    worksheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + row].Style.WrapText = true;

                    worksheet.Cells["M" + row].Value = item.Registro_sanitario;
                    worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + row].Style.WrapText = true;

                    worksheet.Cells["N" + row].Value = item.Bp_manufactura;
                    worksheet.Cells["N" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["N" + row].Style.WrapText = true;

                    worksheet.Cells["O" + row].Value = item.Metodologia_analisis;
                    worksheet.Cells["O" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["O" + row].Style.WrapText = true;

                    worksheet.Cells["P" + row].Value = item.Condicion_almacenamiento;
                    worksheet.Cells["P" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["P" + row].Style.WrapText = true;

                    worksheet.Cells["Q" + row].Value = item.Plazo_entrega;
                    worksheet.Cells["Q" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["Q" + row].Style.WrapText = true;

                    worksheet.Cells["R" + row].Value = item.Precio_unitario;
                    worksheet.Cells["R" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["R" + row].Style.Numberformat.Format = "#,##0.0000";
                    worksheet.Cells["R" + row].Style.WrapText = true;

                    worksheet.Cells["S" + row].Value = item.Total;
                    worksheet.Cells["S" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["S" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["S" + row].Style.WrapText = true;

                    row++;
                }


                worksheet.Cells["R" + row].Value = "TOTAL S/.";
                worksheet.Cells["R" + row].Style.WrapText = true;
                worksheet.Cells["R" + row].Style.Font.Bold = true;
                worksheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["R" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["S" + row].Value = cotizacion.Monto_total;
                worksheet.Cells["S" + row].Style.WrapText = true;
                worksheet.Cells["S" + row].Style.Font.Bold = true;
                worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["S" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row += 7;

                ExcelPicture firmaUnilene = worksheet.Drawings.AddPicture("firma", firma);
                firmaUnilene.SetPosition((row-7), 0, 11, 34);
                firmaUnilene.SetSize(175, 94);

                worksheet.Cells["L" + row + ":O" + row].Merge = true;
                worksheet.Cells["L" + row].Value = "Fima y Sello del Repreentante Legal";
                worksheet.Cells["L" + row].Style.WrapText = true;
                worksheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row++;

                worksheet.Cells["A" + row].Value = "CONDICIONES GENERAL";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;

                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value= "1.- El precio de Mercado será a todo costo, es decir, deberá incluir todos los tributos (incluido el I.G.V.), seguros, transportes, inspecciones";
                worksheet.Cells["A" + row].Style.Font.Size = 11;

                row++;

                worksheet.Row(row).Height= 33.5;
                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "2.- Cabe precisar que frente a un reclamo, queja u observación de los materiales medicos, queda en potestad de Essalud, el sometimiento al control de calidad respectivo, cuyo costo será asumido enteramente por el proveedor, en caso el resultado sea no conforme el plazo para el canje de los saldos inmovilizados en cada alamacen de la entidad es de 15 dias calendarios contados a partir de recibida la comunicación escrita por parte de la entidad.";
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 32;
                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "3.- De haberse efectuado el pago de un lote declarado no conforme, el contratista se obliga a reponer el costo total de las cantidades consumidas y en caso de no efectuarse el canje abonara el costo total del lote inmovilizado a LA ENTIDAD mediante pago en efectivo o cheque de gerencia o deduciendolo de cualqueira de sus facturas";
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Entregas:";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.Font.UnderLine = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Única";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C5D9F1"));

                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Lugar de Entrega:";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.Font.UnderLine = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "La entregas se realizarán en el almacen de la Redes Asistencial Puno, sito en Av. Don Bosco S/N - Urb. Rinconada Salsedo";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C5D9F1"));


                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Plazo de Entrega:";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.Font.UnderLine = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Cells["A" + row + ":S" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "El plazo de entrega es días calendarios según cuadro precedente después de recepcionada la Orden de Compra. Si el plazo propuesto es superior estará sujeto a evaluación por EsSalud.";
                worksheet.Cells["A" + row].Style.Font.Size = 11;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C5D9F1"));

                TextoNegrita(worksheet);
                BordesCeldas(worksheet);
                PintarCeldas(worksheet);

                worksheet.View.ZoomScale = 100;
                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }

            return reporte;
        }

        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1,A2,A15,Q2,R2,F5,Q5,Q6,R6,S6,A17:S24,Q5"].Style.Font.Bold = true;
            worksheet.Cells["F5"].Style.Font.Italic = true;
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 7.57 + 0.71;
            worksheet.Column(2).Width = 10.29 + 0.71;
            worksheet.Column(3).Width = 8 + 0.71;
            worksheet.Column(4).Width = 12.14 + 0.71;
            worksheet.Column(5).Width = 45.43 + 0.71;
            worksheet.Column(6).Width = 6 + 0.71;
            worksheet.Column(7).Width = 10 + 0.71;
            worksheet.Column(8).Width = 13.43 + 0.71;
            worksheet.Column(9).Width = 9.57 + 0.71;
            worksheet.Column(10).Width = 7.29 + 0.71;
            worksheet.Column(11).Width = 9.29 + 0.71;
            worksheet.Column(12).Width = 8.14 + 0.71;
            worksheet.Column(13).Width = 7.14 + 0.71;
            worksheet.Column(14).Width = 7.29 + 0.71;
            worksheet.Column(15).Width = 7.29 + 0.71;
            worksheet.Column(16).Width = 6.43 + 0.71;
            worksheet.Column(17).Width = 8.86 + 0.71;
            worksheet.Column(18).Width = 13.43 + 0.71;
            worksheet.Column(19).Width = 13 + 0.71;

            worksheet.Row(5).Height = 27.75;
            worksheet.Row(6).Height = 39.75;
            worksheet.Row(7).Height = 6;

            worksheet.Row(8).Height = 13.5;
            worksheet.Row(9).Height = 13.5;
            worksheet.Row(10).Height = 13.5;
            worksheet.Row(11).Height = 13.5;
            worksheet.Row(12).Height = 13.5;
            worksheet.Row(13).Height = 13.5;
            worksheet.Row(14).Height = 6.75;
            worksheet.Row(16).Height = 5.25;

            worksheet.Row(17).Height = 25;
            worksheet.Row(18).Height = 33;
            worksheet.Row(19).Height = 24;
            worksheet.Row(20).Height = 14.25;
            worksheet.Row(21).Height = 14.25;
            worksheet.Row(22).Height = 60;
            worksheet.Row(23).Height = 58.5;
            worksheet.Row(24).Height = 84.75;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["Q2:Q3,R2:R3,F5:N6,Q5:S5"].Merge = true;
            worksheet.Cells["A8:D8,A9:D9,A10:D10,A11:D11,A12:D12,A13:D13,L8:M8,L9:M9,L10:M10,L11:M11,L12:M12,L13:M13"].Merge = true;
            worksheet.Cells["E8:H8,E9:H9,E10:H10,E11:H11,E12:H12,E13:H13,N8:S8,N9:S9,N10:S10,N11:S11,N12:S12,N13:S13"].Merge = true;
            worksheet.Cells["A17:A24,B17:C23,D17:G21,D22:D24,E22:E24,F22:F24,G22:G24,H17:P17,H18:H24,I18:I24,J18:J24,K18:P18,K19:L19,M19:P19,K20:K24,L20:L24,M20:M24,N20:N24,O20:O24,P20:P24,Q17:Q24,R17:R24,S17:S24"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["Q2:Q3,R2:R3,F5:N6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells["Q5:S5,Q6,R6,S6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A17:A24,B17:C23,D17:G21,D22:D24,E22:E24,F22:F24,G22:G24,H17:P17,H18:H24,I18:I24,J18:J24,K18:P18,K19:L19,M19:P19,K20:K24,L20:L24,M20:M24,N20:N24,O20:O24,P20:P24,Q17:Q24,R17:R24,S17:S24,B24,C24"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void PintarCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["F5:N6,Q5:S5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["F5:N6,Q5:S5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C0C0C0"));
        }
    }
}
