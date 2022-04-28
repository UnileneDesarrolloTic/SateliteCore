using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public static class Formato5_Report
    {
        public static string Exportar(Image firma,Image logoUnilene, Formato5_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("EsSalud Sabogal");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 50);

                workSheet.Cells.Style.Font.Name = "Calibri Light";
                workSheet.Cells.Style.Font.Size = 9;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);

                workSheet.Cells["A2"].Value = "FORMATO N° 01 - SOLICITUD DE COTIZACIÓN";
                workSheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["A4"].Value = "ESSALUD - OFICINA DE ABASTECIMIENTO Y CONTROL PATRIMONIAL - UNIDAD DE PROGRAMACIÓN";

                workSheet.Cells["U4"].Value = "N° COTIZACIÓN";

                workSheet.Cells["X4"].Value = cotizacion.NroCotizacion;
                workSheet.Cells["X4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["X4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A5"].Value = "RED ASISTENCIAL SABOGAL";

                workSheet.Cells["A7"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A7"].Style.Font.Size = 8;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A8"].Value = "RAZÓN SOCIAL";
                workSheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D8"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["D8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A9"].Value = "DIRECCIÓN";
                workSheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D9"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["D9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A10"].Value = "TELÉFONO(S)";
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D10"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["D10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A11"].Value = "EMAIL";
                workSheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D11"].Value = cotizacion.Prov_Email;
                workSheet.Cells["D11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D11"].Style.Font.UnderLine = true;
                workSheet.Cells["D11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["A12"].Value = "VIGENCIA DE LA COTIZACIÓN";
                workSheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D12"].Value = cotizacion.Prov_Vigencia;
                workSheet.Cells["D12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J7"].Value = "FACTURACIÓN";
                workSheet.Cells["J7"].Style.Font.Size = 8;
                workSheet.Cells["J7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J8"].Value = "RUC";
                workSheet.Cells["J8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["K8"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["K8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J10"].Value = "TELÉFONO AREA CONTABLE";
                workSheet.Cells["J10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["L10"].Value = cotizacion.Cont_Telefono;
                workSheet.Cells["L10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J11"].Value = "EMAIL ÁREA CONTABLE";
                workSheet.Cells["J11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["L11"].Value = cotizacion.Cont_Email;
                workSheet.Cells["L11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J12"].Value = "CONTACTO ÁREA CONTABLE";
                workSheet.Cells["J12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["L12"].Value = cotizacion.Cont_Area;
                workSheet.Cells["L12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["Q7"].Value = "DATOS DE CONTACTO - REPRESENTANTE DE VENTAS";
                workSheet.Cells["Q7"].Style.Font.Size = 8;
                workSheet.Cells["Q7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["Q8"].Value = "APELLIDOS Y NOMBRES";
                workSheet.Cells["Q8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["T8"].Value = cotizacion.Ctac_Nombres;
                workSheet.Cells["T8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["Q9"].Value = "CARGO";
                workSheet.Cells["Q9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["T9"].Value = cotizacion.Ctac_Cargo;
                workSheet.Cells["T9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["Q10"].Value = "N° CELULAR";
                workSheet.Cells["Q10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["T10"].Value = cotizacion.Ctac_Celular;
                workSheet.Cells["T10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["Q11"].Value = "EMAIL N°1";
                workSheet.Cells["Q11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["T11"].Value = cotizacion.Ctac_Email1;
                workSheet.Cells["T11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T11"].Style.Font.UnderLine = true;
                workSheet.Cells["T11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["Q12"].Value = "EMAIL N°2";
                workSheet.Cells["Q12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["T12"].Value = cotizacion.Ctac_Email2;
                workSheet.Cells["T12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T12"].Style.Font.UnderLine = true;
                workSheet.Cells["T12"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["A14"].Value = "N° Item";
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["B14"].Value = "DESCRIPCIÓN DEL REQUERIMIENTO";
                workSheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F14"].Value = "DESCRIPCIÓN DEL PRODUCTO OFERTADO";
                workSheet.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["K14"].Value = "REQUISITOS TÉCNICOS DEL POSTOR";
                workSheet.Cells["K14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K14"].Style.WrapText = true;

                workSheet.Cells["M14"].Value = "REQUISITOS TÉCNICOS DEL DISPOSITIVO: DEL DISPOSITIVO MÉDICO";
                workSheet.Cells["M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["X14"].Value = "DEL PRECIO";
                workSheet.Cells["X14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["X14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["B15"].Value = "Código SAP";
                workSheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C15"].Value = "Denominación";
                workSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D15"].Value = "UM";
                workSheet.Cells["D15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["E15"].Value = "Cantidad";
                workSheet.Cells["E15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F15"].Value = "Marca / Modelo";
                workSheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F15"].Style.WrapText = true;

                workSheet.Cells["G15"].Value = "Procedencia";
                workSheet.Cells["G15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G15"].Style.TextRotation = 90;

                workSheet.Cells["H15"].Value = "Forma de Presentación";
                workSheet.Cells["H15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H15"].Style.TextRotation = 90;

                workSheet.Cells["I15"].Value = "Plazo de entrega (*) indicar días calendarios o hábiles (en caso de no precisar, se considerada como días calendarios)";
                workSheet.Cells["I15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I15"].Style.WrapText = true;

                workSheet.Cells["J15"].Value = "Vigencia mínima  del Producto deberá ser igual o mayor a dieciocho (18) meses";
                workSheet.Cells["J15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J15"].Style.WrapText = true;

                workSheet.Cells["K15"].Value = "Resolución de Autorización Sanitaria de Funcionamiento de Establecimiento Farmacéutico vigente";
                workSheet.Cells["K15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K15"].Style.TextRotation = 90;
                workSheet.Cells["K15"].Style.WrapText = true;

                workSheet.Cells["L15"].Value = "Certificado de Buenas Prácticas de Almacenamiento (CBPA)";
                workSheet.Cells["L15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L15"].Style.TextRotation = 90;
                workSheet.Cells["L15"].Style.WrapText = true;

                workSheet.Cells["M15"].Value = "Registro Sanitario o Certificado de Registro Sanitario";
                workSheet.Cells["M15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M15"].Style.TextRotation = 90;
                workSheet.Cells["M15"].Style.WrapText = true;

                workSheet.Cells["N15"].Value = "Carta de Representación";
                workSheet.Cells["N15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N15"].Style.TextRotation = 90;

                workSheet.Cells["O15"].Value = "Certificado de Buenas Prácticas de Manufactura (CBPM)";
                workSheet.Cells["O15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O15"].Style.TextRotation = 90;
                workSheet.Cells["O15"].Style.WrapText = true;

                workSheet.Cells["P15"].Value = "Certificado de Análisis del producto farmacéutico terminado (Protocolo de Análisis)";
                workSheet.Cells["P15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P15"].Style.TextRotation = 90;
                workSheet.Cells["P15"].Style.WrapText = true;

                workSheet.Cells["Q15"].Value = "Metodología de Análisis (Copia Simple)";
                workSheet.Cells["Q15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q15"].Style.TextRotation = 90;

                workSheet.Cells["R15"].Value = "Ficha Técnica del producto (Copia simple)";
                workSheet.Cells["R15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R15"].Style.TextRotation = 90;

                workSheet.Cells["S15"].Value = "Folletería / Manual de Instrucciones de Uso o Inserto (original o copia simple)";
                workSheet.Cells["S15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S15"].Style.TextRotation = 90;
                workSheet.Cells["S15"].Style.WrapText = true;

                workSheet.Cells["T15"].Value = "Declaración Jurada de Presentación del dispositivo médico ofertado, de compromiso del plazo de entrega y vigencia";
                workSheet.Cells["T15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T15"].Style.TextRotation = 90;
                workSheet.Cells["T15"].Style.WrapText = true;

                workSheet.Cells["U15"].Value = "Declaración Jurada de Compromiso de Canje y/o Reposición por Defectos o Vicios Ocultos";
                workSheet.Cells["U15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["U15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["U15"].Style.TextRotation = 90;
                workSheet.Cells["U15"].Style.WrapText = true;

                workSheet.Cells["V15"].Value = "Muestra ( SI / NO )";
                workSheet.Cells["V15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["V15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V15"].Style.TextRotation = 90;

                workSheet.Cells["W15"].Value = "La Oferta incluye rotulado: ESSALUD PROHIBIDA SU VENTA ( SI / NO )";
                workSheet.Cells["W15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["W15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["W15"].Style.TextRotation = 90;
                workSheet.Cells["W15"].Style.WrapText = true;

                workSheet.Cells["X15"].Value = "PRECIO UNITARIO A OFERTAR S/. Inc. IGV";
                workSheet.Cells["X15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["X15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["X15"].Style.TextRotation = 90;
                workSheet.Cells["X15"].Style.WrapText = true;

                workSheet.Cells["Y15"].Value = "VALOR TOTAL DEL ÍTEM S /.";
                workSheet.Cells["Y15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Y15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Y15"].Style.TextRotation = 90;

                int index = 16;

                foreach (Formato5_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Cells["A" + index].Value = index - 15;
                    workSheet.Cells["A" + index].Style.Font.Size = 11;
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                    workSheet.Cells["B" + index].Value = item.CodigoSap;
                    workSheet.Cells["B" + index].Style.Font.Size = 11;
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;

                    workSheet.Cells["C" + index].Value = item.Denominacion;
                    workSheet.Cells["C" + index].Style.Font.Size = 11;
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;

                    workSheet.Cells["D" + index].Value = item.Um;
                    workSheet.Cells["D" + index].Style.Font.Size = 11;
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = item.Cantidad;
                    workSheet.Cells["E" + index].Style.Font.Size = 11;
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["F" + index].Value = item.Marca;
                    workSheet.Cells["F" + index].Style.Font.Size = 11;
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;

                    workSheet.Cells["G" + index].Value = item.Procedencia;
                    workSheet.Cells["G" + index].Style.Font.Size = 11;
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;

                    workSheet.Cells["H" + index].Value = item.Presentacion;
                    workSheet.Cells["H" + index].Style.Font.Size = 11;
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;

                    workSheet.Cells["I" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["I" + index].Style.Font.Size = 11;
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Cells["J" + index].Value = item.VigenciaMinima;
                    workSheet.Cells["J" + index].Style.Font.Size = 11;
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;

                    workSheet.Cells["K" + index].Value = item.AutoSanitario;
                    workSheet.Cells["K" + index].Style.Font.Size = 11;
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + index].Value = item.BuenasPracAlmace;
                    workSheet.Cells["L" + index].Style.Font.Size = 11;
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + index].Value = item.RegistroSanitario;
                    workSheet.Cells["M" + index].Style.Font.Size = 11;
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["N" + index].Value = item.CartaRepresentacion;
                    workSheet.Cells["N" + index].Style.Font.Size = 11;
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;

                    workSheet.Cells["O" + index].Value = item.Cbpm;
                    workSheet.Cells["O" + index].Style.Font.Size = 11;
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + index].Value = item.CertificadoAnalisis;
                    workSheet.Cells["P" + index].Style.Font.Size = 11;
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["Q" + index].Value = item.MetodologiaAnalisis;
                    workSheet.Cells["Q" + index].Style.Font.Size = 11;
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["R" + index].Value = item.FichaTecnica;
                    workSheet.Cells["R" + index].Style.Font.Size = 11;
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["S" + index].Value = item.Folleteria;
                    workSheet.Cells["S" + index].Style.Font.Size = 11;
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["T" + index].Value = item.CompPlazoEntrega;
                    workSheet.Cells["T" + index].Style.Font.Size = 11;
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["U" + index].Value = item.CompReposicionDefecto;
                    workSheet.Cells["U" + index].Style.Font.Size = 11;
                    workSheet.Cells["U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["V" + index].Value = item.Muestra;
                    workSheet.Cells["V" + index].Style.Font.Size = 11;
                    workSheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["W" + index].Value = item.RotuladoESSALUD;
                    workSheet.Cells["W" + index].Style.Font.Size = 11;
                    workSheet.Cells["W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["X" + index].Value = item.PrecioUnitario;
                    workSheet.Cells["X" + index].Style.Font.Size = 11;
                    workSheet.Cells["X" + index].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["X" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["X" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells["X" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["Y" + index].Value = item.PrecioTotal;
                    workSheet.Cells["Y" + index].Style.Font.Size = 11;
                    workSheet.Cells["Y" + index].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["Y" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Y" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells["Y" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    index++;
                }

                workSheet.Cells["A" + index].Value = "CONSIDERACIONES GENERALES";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Bold = true;

                workSheet.Cells["W" + index + ":X" + index].Merge = true;
                workSheet.Cells["W" + index].Value = "Total S/.";
                workSheet.Cells["W" + index].Style.Font.Bold = true;
                workSheet.Cells["W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["W" + index + ":X" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["Y" + index].Value = cotizacion.Foot_Total;
                workSheet.Cells["Y" + index].Style.Font.Bold = true;
                workSheet.Cells["Y" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["Y" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Y" + index].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["Y" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index = index + 2;

                workSheet.Row(index).Height = 26;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "1.- El precio de Mercado será a todo costo, es decir, deberá incluir todos los tributos (incluido el I.G.V.), seguros, " +
                    "transportes, inspecciones, pruebas, costo total de los controles de calidad (financiamiento al 100%) y cualquier otro concepto que pueda tener incidencia " +
                    "sobre el costo del bien, hasta la entrega en su destino.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "2.- El Valor Referencial Total debe ser expresado como maximo en 2 (DOS) decimales, el precio unitario se podrá expresar hasta en 3 (TRES) decimales.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Row(index).Height = 26;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "3.- En el caso que la empresa postora sea a la vez la empresa fabricante nacional; en mérito a la aplicación de los dispositivos " +
                    "que en esta materia se encuentran vigentes en el territorio peruano, y de tener Certificado de Buenas Practicas de Manufactura (CBPM) indicar \"si\", en la columna " +
                    "de Certificado de Buenas Practicas de Almacenamiento (CBPA) indicar \"si\" entendiendose que el CBPM contiene al CBPA. Asimismo, podrá indicar \"NO\" en la columa " +
                    "de Carta de Representación, pero debera de indicar que es FABRICANTE.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "4.- Documentos alternativos del CBPM: Ó Certificado de Libre Venta ó Certificado de Libre Comercialización ó Certificado del Producto.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "5.- Asimismo, en la Columna de Carta de Representante de ser FABRICANTE escribir la palabra; y de no serlo (Distribuidor Autorizado) indicarlo también.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "6.- DE SER EL CASO INDICAR SI SU PRECIO DEL MATERIAL O INSUMO ESTA EXONERADO DEL IGV Y ARANCELES.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "7.- Otras consideraciones se encuentran descritas en el requerimiento del area usuaria.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.WrapText = true;

                ExcelPicture firmaCatizacion = workSheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(index+1, -8, 1, 85);
                firmaCatizacion.SetSize(265, 120);

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
            workSheet.Column(1).Width = 6.17 + 0.71;
            workSheet.Column(2).Width = 9.67 + 0.71;
            workSheet.Column(3).Width = 29.83 + 0.71;
            workSheet.Column(4).Width = 4.83 + 0.71;
            workSheet.Column(5).Width = 11 + 0.71;
            workSheet.Column(6).Width = 11.17 + 0.71;
            workSheet.Column(7).Width = 8.83 + 0.71;
            workSheet.Column(8).Width = 10.5 + 0.71;
            workSheet.Column(9).Width = 11.33 + 0.71;
            workSheet.Column(10).Width = 11.17 + 0.71;
            workSheet.Column(11).Width = 7.83 + 0.71;
            workSheet.Column(12).Width = 6.83 + 0.71;
            workSheet.Column(13).Width = 5.5 + 0.71;
            workSheet.Column(14).Width = 11.67 + 0.71;
            workSheet.Column(15).Width = 6.67 + 0.71;
            workSheet.Column(16).Width = 6.67 + 0.71;
            workSheet.Column(17).Width = 6.67 + 0.71;
            workSheet.Column(18).Width = 6.67 + 0.71;
            workSheet.Column(19).Width = 6.67 + 0.71;
            workSheet.Column(20).Width = 6.67 + 0.71;
            workSheet.Column(21).Width = 6.67 + 0.71;
            workSheet.Column(22).Width = 6.67 + 0.71;
            workSheet.Column(23).Width = 6.67 + 0.71;
            workSheet.Column(24).Width = 11.17 + 0.71;
            workSheet.Column(25).Width = 15.67 + 0.71;

            workSheet.Row(7).Height = 27;
            workSheet.Row(8).Height = 25.5;
            workSheet.Row(9).Height = 25.5;
            workSheet.Row(10).Height = 25.5;
            workSheet.Row(11).Height = 25.5;
            workSheet.Row(12).Height = 25.5;
            workSheet.Row(14).Height = 38;
            workSheet.Row(15).Height = 186;
        }
        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A2:Y2"].Merge = true;
            workSheet.Cells["A4:G4"].Merge = true;
            workSheet.Cells["A5:C5"].Merge = true;
            workSheet.Cells["X4:Y4"].Merge = true;
            workSheet.Cells["A7:H7"].Merge = true;
            workSheet.Cells["A8:C8"].Merge = true;
            workSheet.Cells["A9:C9"].Merge = true;
            workSheet.Cells["U4:V4"].Merge = true;
            workSheet.Cells["A10:C10"].Merge = true;
            workSheet.Cells["A11:C11"].Merge = true;
            workSheet.Cells["A12:C12"].Merge = true;
            workSheet.Cells["D8:H8"].Merge = true;
            workSheet.Cells["D9:H9"].Merge = true;
            workSheet.Cells["D10:H10"].Merge = true;
            workSheet.Cells["D11:H11"].Merge = true;
            workSheet.Cells["D12:H12"].Merge = true;
            workSheet.Cells["J7:O7"].Merge = true;
            workSheet.Cells["K8:O8"].Merge = true;
            workSheet.Cells["J10:K10"].Merge = true;
            workSheet.Cells["J11:K11"].Merge = true;
            workSheet.Cells["J12:K12"].Merge = true;
            workSheet.Cells["L10:O10"].Merge = true;
            workSheet.Cells["L11:O11"].Merge = true;
            workSheet.Cells["L12:O12"].Merge = true;
            workSheet.Cells["Q7:Y7"].Merge = true;
            workSheet.Cells["Q8:S8,T8:Y8"].Merge = true;
            workSheet.Cells["Q9:S9,T9:Y9"].Merge = true;
            workSheet.Cells["Q10:S10,T10:Y10"].Merge = true;
            workSheet.Cells["Q11:S11,T11:Y11"].Merge = true;
            workSheet.Cells["Q12:S12,T12:Y12"].Merge = true;
            workSheet.Cells["A14:A15,B14:E14,F14:J14,K14:L14,M14:W14,X14:Y14"].Merge = true;
        }
        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A7:H12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A8:C8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A9:C9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A10:C10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A11:C11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A12:C12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D8:H8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D9:H9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D10:H10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D11:H11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D12:H12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["X4:Y4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["J7:O7,J8,K8:O8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["J10:K10,J11:K11,J12:K12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L10:O10,L11:O11,L12:O12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q7:Y7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q8:S8,T8:Y8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q9:S9,T9:Y9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q10:S10,T10:Y10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q11:S11,T11:Y11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q12:S12,T12:Y12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A14:A15,B14:E14,F14:J14,K14:L14,M14:W14,X14:Y14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B15,C15,D15,E15,F15,G15,H15,I15,J15,K15,L15,M15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N15,O15,P15,Q15,R15,S15,T15,U15,V15,W15,X15,Y15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A7,A8:A12,J7:O7,J8,J10:J12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["A7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BC2E6"));
            workSheet.Cells["A8:A12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));
            workSheet.Cells["J7:O7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F8CBAD"));
            workSheet.Cells["J8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FCE4D6"));
            workSheet.Cells["J10:J12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FCE4D6"));

            workSheet.Cells["Q7:Y7,Q8:S8,Q9:S9,Q10:S10,Q11:S11,Q12:S12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["Q7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFE699"));
            workSheet.Cells["Q8:S12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"));

            workSheet.Cells["A14:A15,B14:E14,B15:E15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["A14:A15,B14:E14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BC2E6"));
            workSheet.Cells["B15:E15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

            workSheet.Cells["F14:Y15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["F14:J14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F8CBAD"));
            workSheet.Cells["K14:L14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFE699"));
            workSheet.Cells["M14:W14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B4C6E7"));
            workSheet.Cells["X14:Y14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C6E0B4"));

            workSheet.Cells["F15:J15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FCE4D6"));
            workSheet.Cells["K15:L15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"));
            workSheet.Cells["M15:W15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9E1F2"));
            workSheet.Cells["X15:Y15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E2EFDA"));

        }
        private static void TextoNegrita(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A2"].Style.Font.Bold = true;
            workSheet.Cells["U4:Y4"].Style.Font.Bold = true;
            workSheet.Cells["A7:Y7"].Style.Font.Bold = true;
            workSheet.Cells["A7"].Style.Font.Bold = true;
            workSheet.Cells["X4"].Style.Font.Bold = true;
            workSheet.Cells["Q7"].Style.Font.Bold = true;
            workSheet.Cells["J7"].Style.Font.Bold = true;
            workSheet.Cells["A14"].Style.Font.Bold = true;
            workSheet.Cells["B14:Y14"].Style.Font.Bold = true;
        }

    }
}
