using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato3_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato3_Model cotizacion)
        {
            byte[] file;

            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Almenara");

                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 5, 0, 9);
                imagenUnilene.SetSize(197, 99);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Name = "Calibri";

                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);
                PintarCeldas(workSheet);

                workSheet.Cells["C2"].Value = "UNIDAD DE PROGRAMACIÓN";
                workSheet.Cells["C2"].Style.Font.Size = 11;
                workSheet.Cells["C2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C3"].Value = "OFICINA DE ABASTECIMIENTO Y CONTROL PATRIMONIAL";
                workSheet.Cells["C3"].Style.Font.Size = 11;
                workSheet.Cells["C3"].Style.WrapText = true;
                workSheet.Cells["C3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C4"].Value = "RED ASISTENCIAL ALMENARA";
                workSheet.Cells["C4"].Style.Font.Size = 12;
                workSheet.Cells["C4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                
                workSheet.Cells["S2"].Value = "FORMATO DE COTIZACIÓN";
                workSheet.Cells["S2"].Style.Font.Size = 18;

                workSheet.Cells["S3"].Value = "Versión: 01-2017";
                workSheet.Cells["S3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S3"].Style.Font.Size = 18;

                workSheet.Cells["S4"].Value = "Página:  1 / 1";
                workSheet.Cells["S4"].Style.Font.Size = 18;


                workSheet.Cells["A7"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A7"].Style.Font.Size = 20;
                workSheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["A8"].Value = "RAZÓN SOCIAL";
                workSheet.Cells["A8"].Style.Font.Size = 20;
                workSheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A9"].Value = "RUC";
                workSheet.Cells["A9"].Style.Font.Size = 20;
                workSheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A10"].Value = "DIRECCIÓN";
                workSheet.Cells["A10"].Style.Font.Size = 20;
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A11"].Value = "E-MAIL";
                workSheet.Cells["A11"].Style.Font.Size = 20;
                workSheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A12"].Value = "N° COTIZACIÓN";
                workSheet.Cells["A12"].Style.Font.Size = 20;
                workSheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C8"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C8"].Style.Font.Size = 20;
                workSheet.Cells["C8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C9"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["C9"].Style.Font.Size = 20;
                workSheet.Cells["C9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C10"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["C10"].Style.Font.Size = 20;
                workSheet.Cells["C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C11"].Value = cotizacion.Prov_Correo;
                workSheet.Cells["C11"].Style.Font.Size = 16;
                workSheet.Cells["C11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C11"].Style.Font.Name = "Arial";
                workSheet.Cells["C11"].Style.Font.UnderLine = true;
                workSheet.Cells["C11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["C12"].Value = cotizacion.Prov_NroCotizacion;
                workSheet.Cells["C12"].Style.Font.Size = 20;
                workSheet.Cells["C12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D8"].Value = "TELÉFONO";
                workSheet.Cells["D8"].Style.Font.Size = 20;
                workSheet.Cells["D8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D9"].Value = "FAX";
                workSheet.Cells["D9"].Style.Font.Size = 20;
                workSheet.Cells["D9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D10"].Value = "VIGENCIA DE OFERTA";
                workSheet.Cells["D10"].Style.Font.Size = 20;
                workSheet.Cells["D10"].Style.WrapText = true;
                workSheet.Cells["D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D11"].Value = "FECHA";
                workSheet.Cells["D11"].Style.Font.Size = 20;
                workSheet.Cells["D11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D12"].Value = "DATOS ADICIONALES";
                workSheet.Cells["D12"].Style.Font.Size = 16;
                workSheet.Cells["D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F8"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["F8"].Style.Font.Size = 20;
                workSheet.Cells["F8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F9"].Value = cotizacion.Prov_Fax;
                workSheet.Cells["F9"].Style.Font.Size = 20;
                workSheet.Cells["F9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F10"].Value = cotizacion.Prov_Vigencia;
                workSheet.Cells["F10"].Style.Font.Size = 20;
                workSheet.Cells["F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F11"].Value = cotizacion.Prov_Fecha;
                workSheet.Cells["F11"].Style.Font.Size = 20;
                workSheet.Cells["F11"].Style.Numberformat.Format = "dd/MM/yyyy";
                workSheet.Cells["F11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["G12"].Value = cotizacion.Prov_DatosAdicionales;
                workSheet.Cells["G12"].Style.Font.Size = 16;
                workSheet.Cells["G12"].Style.Font.Name = "Arial";
                workSheet.Cells["G12"].Style.Font.UnderLine = true;
                workSheet.Cells["G12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G12"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["H8"].Value = "CONTACTO";
                workSheet.Cells["H8"].Style.Font.Size = 20;
                workSheet.Cells["H8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["H9"].Value = "CARGO";
                workSheet.Cells["H9"].Style.Font.Size = 20;
                workSheet.Cells["H9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["H10"].Value = "TELÉFONO";
                workSheet.Cells["H10"].Style.Font.Size = 20;
                workSheet.Cells["H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["H11"].Value = "E-MAIL";
                workSheet.Cells["H11"].Style.Font.Size = 20;
                workSheet.Cells["H11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                workSheet.Cells["J8"].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["J8"].Style.Font.Size = 20;
                workSheet.Cells["J8"].Style.WrapText = true;
                workSheet.Cells["J8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J9"].Value = cotizacion.Prov_Cargo;
                workSheet.Cells["J9"].Style.WrapText = true;
                workSheet.Cells["J9"].Style.Font.Size = 20;
                workSheet.Cells["J9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J10"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["J10"].Style.Font.Size = 19;
                workSheet.Cells["J10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["J11"].Value = cotizacion.Prov_Email;
                workSheet.Cells["J11"].Style.Font.Size = 14;
                workSheet.Cells["J11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J11"].Style.Font.Name = "Arial";
                workSheet.Cells["J11"].Style.Font.UnderLine = true;
                workSheet.Cells["J11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["N7"].Value = "DATOS DEL AREA SOLICITANTE";
                workSheet.Cells["N7"].Style.Font.Size = 20;
                workSheet.Cells["N7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["N8"].Value = "ÁREA USUARIA";
                workSheet.Cells["N8"].Style.Font.Size = 20;
                workSheet.Cells["N8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["N9"].Value = "DEPENDENCIA";
                workSheet.Cells["N9"].Style.Font.Size = 20;
                workSheet.Cells["N9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["N10"].Value = "CONTACTO - PROGRAMACIÓN";
                workSheet.Cells["N10"].Style.Font.Size = 20;
                workSheet.Cells["N10"].Style.WrapText = true;
                workSheet.Cells["N10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["N11"].Value = "E-MAIL";
                workSheet.Cells["N11"].Style.Font.Size = 20;
                workSheet.Cells["N11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["N12"].Value = "TELÉFONO";
                workSheet.Cells["N12"].Style.Font.Size = 20;
                workSheet.Cells["N12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["P8"].Value = cotizacion.Soli_AreaUsuario;
                workSheet.Cells["P8"].Style.Font.Size = 18;
                workSheet.Cells["P8"].Style.WrapText = true;
                workSheet.Cells["P8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["P9"].Value = cotizacion.Soli_Dependencia;
                workSheet.Cells["P9"].Style.Font.Size = 15;
                workSheet.Cells["P9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["P10"].Value = cotizacion.Soli_Contacto;
                workSheet.Cells["P10"].Style.Font.Size = 17;
                workSheet.Cells["P10"].Style.WrapText = true;
                workSheet.Cells["P10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["P11"].Value = cotizacion.Soli_Correo;
                workSheet.Cells["P11"].Style.Font.Size = 14;
                workSheet.Cells["P11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P11"].Style.Font.UnderLine = true;
                workSheet.Cells["P11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));
                workSheet.Cells["P11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["P12"].Value = cotizacion.Soli_Telefono;
                workSheet.Cells["P12"].Style.Font.Size = 14;
                workSheet.Cells["P12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // DETALLE

                workSheet.Cells["A15"].Value = "N° ITEM";
                workSheet.Cells["A15"].Style.Font.Size = 18;
                workSheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A15"].Style.TextRotation = 90;

                workSheet.Cells["B15"].Value = "DESCRIPCIÓN DEL ÍTEM";
                workSheet.Cells["B15"].Style.Font.Size = 18;
                workSheet.Cells["B15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["B17"].Value = "CÓDIGO SAP";
                workSheet.Cells["B17"].Style.Font.Size = 18;
                workSheet.Cells["B17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["C17"].Value = "DENOMINACIÓN";
                workSheet.Cells["C17"].Style.Font.Size = 18;
                workSheet.Cells["C17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                
                workSheet.Cells["D17"].Value = "Unidad de Medida";
                workSheet.Cells["D17"].Style.Font.Size = 18;
                workSheet.Cells["D17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D17"].Style.TextRotation = 90;

                workSheet.Cells["E15"].Value = "CANT. TOTAL";
                workSheet.Cells["E15"].Style.Font.Size = 18;
                workSheet.Cells["E15"].Style.WrapText = true;
                workSheet.Cells["E15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["F15"].Value = "REQUERIMIENTOS MÍNIMOS";
                workSheet.Cells["F15"].Style.Font.Size = 18;
                workSheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["F16"].Value = "Marca";
                workSheet.Cells["F16"].Style.Font.Size = 18;
                workSheet.Cells["F16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F16"].Style.TextRotation = 90;

                workSheet.Cells["F16"].Value = "Marca";
                workSheet.Cells["F16"].Style.Font.Size = 18;
                workSheet.Cells["F16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F16"].Style.TextRotation = 90;

                workSheet.Cells["G16"].Value = "País de Procedencia del Producto";
                workSheet.Cells["G16"].Style.Font.Size = 18;
                workSheet.Cells["G16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G16"].Style.TextRotation = 90;

                workSheet.Cells["H16"].Value = "Forma de Presentación";
                workSheet.Cells["H16"].Style.Font.Size = 18;
                workSheet.Cells["H16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H16"].Style.TextRotation = 90;

                workSheet.Cells["I16"].Value = "Modelo";
                workSheet.Cells["I16"].Style.Font.Size = 18;
                workSheet.Cells["I16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I16"].Style.TextRotation = 90;

                workSheet.Cells["J16"].Value = "RNP Vigente a la fecha  (SI-NO)";
                workSheet.Cells["J16"].Style.Font.Size = 18;
                workSheet.Cells["J16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J16"].Style.TextRotation = 90;

                workSheet.Cells["K16"].Value = "CALIDAD";
                workSheet.Cells["K16"].Style.Font.Size = 18;
                workSheet.Cells["K16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["K17"].Value = "La vigencia mínima del material medico deberá ser de 24 meses contados a partir de la entrega. (salvo excepción) (Si o No)";
                workSheet.Cells["K17"].Style.Font.Size = 18;
                workSheet.Cells["K17"].Style.WrapText = true;
                workSheet.Cells["K17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K17"].Style.TextRotation = 90;

                workSheet.Cells["L17"].Value = "Cumple al 100% con la denominación y descripción del ítem  (Si  ó No)";
                workSheet.Cells["L17"].Style.Font.Size = 18;
                workSheet.Cells["L17"].Style.WrapText = true;
                workSheet.Cells["L17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L17"].Style.TextRotation = 90;

                workSheet.Cells["M17"].Value = "Carta de Representación (En idioma castellano y en original o copia simple). (fecha de emisión no mayor a 2 años), a su nombre y expedida por el fabricante o dueño de marca.De ser fabricante indicarlo (Si o No)";
                workSheet.Cells["M17"].Style.Font.Size = 18;
                workSheet.Cells["M17"].Style.WrapText = true;
                workSheet.Cells["M17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M17"].Style.TextRotation = 90;

                workSheet.Cells["N17"].Value = "Cuenta con Registro Sanitario O Certificado de Registro Sanitario vigente (Si o No)";
                workSheet.Cells["N17"].Style.Font.Size = 18;
                workSheet.Cells["N17"].Style.WrapText = true;
                workSheet.Cells["N17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N17"].Style.TextRotation = 90;

                workSheet.Cells["O17"].Value = "Protocolo y/o Certificado de Análisis (En idioma castellano y en original o copia simple). en concordancia al D.S.010-2001-SA y su modificatoria (Si o No)";
                workSheet.Cells["O17"].Style.Font.Size = 18;
                workSheet.Cells["O17"].Style.WrapText = true;
                workSheet.Cells["O17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O17"].Style.TextRotation = 90;

                workSheet.Cells["P17"].Value = "Cumple al 100% con la Especificación Técnica (Si  ó No)";
                workSheet.Cells["P17"].Style.Font.Size = 18;
                workSheet.Cells["P17"].Style.WrapText = true;
                workSheet.Cells["P17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P17"].Style.TextRotation = 90;

                workSheet.Cells["Q17"].Value = "Certificado de Buena Práctica de Almacenamiento – (BPA) vigente (Si o No)";
                workSheet.Cells["Q17"].Style.Font.Size = 18;
                workSheet.Cells["Q17"].Style.WrapText = true;
                workSheet.Cells["Q17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q17"].Style.TextRotation = 90;
                
                workSheet.Cells["R17"].Value = "Certificado de Buena Práctica de Manufactura del laboratorio fabricante – (BPM) vigente.(SI  ó  NO)";
                workSheet.Cells["R17"].Style.Font.Size = 18;
                workSheet.Cells["R17"].Style.WrapText = true;
                workSheet.Cells["R17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R17"].Style.TextRotation = 90;

                workSheet.Cells["S16"].Value = "Cant.";
                workSheet.Cells["S16"].Style.Font.Size = 18;
                workSheet.Cells["S16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["S17"].Value = "Capacidad de Atención al 100% de lo solicitado (Si o No)";
                workSheet.Cells["S17"].Style.Font.Size = 18;
                workSheet.Cells["S17"].Style.WrapText = true;
                workSheet.Cells["S17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S17"].Style.TextRotation = 90;

                workSheet.Cells["T15"].Value = "Plazo de Entrega en Días Calendarios";
                workSheet.Cells["T15"].Style.Font.Size = 18;
                workSheet.Cells["T15"].Style.WrapText = true;
                workSheet.Cells["T15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T15"].Style.TextRotation = 90;

                workSheet.Cells["U15"].Value = "PRECIO UNITARIO A OFERTAR \nS /.";
                workSheet.Cells["U15"].Style.Font.Size = 18;
                workSheet.Cells["U15"].Style.WrapText = true;
                workSheet.Cells["U15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["U15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["U15"].Style.TextRotation = 90;

                workSheet.Cells["V15"].Value = "VALOR TOTAL A OFERTAR \nS /.";
                workSheet.Cells["V15"].Style.Font.Size = 18;
                workSheet.Cells["V15"].Style.WrapText = true;
                workSheet.Cells["V15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["V15"].Style.TextRotation = 90;

                int index = 20;

                foreach (Coti_Formato3_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Cells["A" + index].Value = index - 19;
                    workSheet.Cells["A" + index].Style.Font.Size = 16;
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                    workSheet.Cells["B" + index].Value = item.CodigoSAP;
                    workSheet.Cells["B" + index].Style.Font.Size = 16;
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = item.Denominacion;
                    workSheet.Cells["C" + index].Style.Font.Size = 16;
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.WrapText = true;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = item.UndMedida;
                    workSheet.Cells["D" + index].Style.Font.Size = 18;
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.TextRotation = 90;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = item.CantidadTotal;
                    workSheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells["E" + index].Style.Font.Size = 16;
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = item.Marca;
                    workSheet.Cells["F" + index].Style.Font.Size = 16;
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.TextRotation = 90;
                    workSheet.Cells["F" + index].Style.WrapText = true;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = item.PaisProcedencia;
                    workSheet.Cells["G" + index].Style.Font.Size = 16;
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.TextRotation = 90;
                    workSheet.Cells["G" + index].Style.WrapText = true;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = item.Presentacion;
                    workSheet.Cells["H" + index].Style.Font.Size = 16;
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.WrapText = true;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = item.Modelo;
                    workSheet.Cells["I" + index].Style.Font.Size = 13;
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.TextRotation = 90;
                    workSheet.Cells["I" + index].Style.WrapText = true;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = item.RnpVigente;
                    workSheet.Cells["J" + index].Style.Font.Size = 18;
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = item.VigenciaMinima;
                    workSheet.Cells["K" + index].Style.Font.Size = 18;
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.WrapText = true;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + index].Value = item.CumpleDenominacionItem;
                    workSheet.Cells["L" + index].Style.Font.Size = 18;
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + index].Value = item.CartaPresentacion;
                    workSheet.Cells["M" + index].Style.Font.Size = 18;
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["N" + index].Value = item.RegistroSanitario;
                    workSheet.Cells["N" + index].Style.Font.Size = 18;
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["O" + index].Value = item.CertificadoAnalisis;
                    workSheet.Cells["O" + index].Style.Font.Size = 18;
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + index].Value = item.CumpleEspecificacionTec;
                    workSheet.Cells["P" + index].Style.Font.Size = 18;
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["Q" + index].Value = item.PracticaAlmacenamiento;
                    workSheet.Cells["Q" + index].Style.Font.Size = 18;
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["R" + index].Value = item.PracticaManufactura;
                    workSheet.Cells["R" + index].Style.Font.Size = 18;
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["S" + index].Value = item.CapacidadAtencion;
                    workSheet.Cells["S" + index].Style.Font.Size = 18;
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["T" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["T" + index].Style.Font.Size = 16;
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.TextRotation = 90;
                    workSheet.Cells["T" + index].Style.WrapText = true;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["U" + index].Value = item.PrecioUnitario;
                    workSheet.Cells["U" + index].Style.Font.Size = 16;
                    workSheet.Cells["U" + index].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["V" + index].Value = item.ValorTotal;
                    workSheet.Cells["V" + index].Style.Font.Size = 16;
                    workSheet.Cells["V" + index].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    index++;
                }


                workSheet.Cells["T" + index + ":U" + index].Merge = true;
                workSheet.Cells["T" + index + ":U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["T" + index].Value = "TOTAL S/.";
                workSheet.Cells["T" + index].Style.Font.Size = 20;
                workSheet.Cells["T" + index].Style.Font.Bold = true;
                workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["V" + index].Value = cotizacion.Total;
                workSheet.Cells["V" + index].Style.Font.Size = 20;
                workSheet.Cells["V" + index].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["V" + index].Style.Font.Bold = true;
                workSheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                index++;

                workSheet.Cells["A" + index + ":C" + index].Merge = true;

                workSheet.Cells["A" + index].Value = "FORMA DE PAGO:  " + cotizacion.Foot_FormaPago;
                workSheet.Cells["A" + index].Style.Font.Size = 19;
                workSheet.Cells["A" + index].Style.Font.Color.SetColor(Color.Red);
                workSheet.Cells["A" + index].Style.Font.Bold = true;

                index++;

                workSheet.Cells["A" + index].Value = "Observaciones";
                workSheet.Cells["A" + index].Style.Font.Size = 16;
                workSheet.Cells["A" + index].Style.Font.Bold = true;

                index++;

                workSheet.Cells["A" + index].Value = cotizacion.Foot_Observaciones;
                workSheet.Cells["A" + index].Style.Font.Size = 20;
                workSheet.Cells["A" + index].Style.Font.Bold = true;

                index = index + 2;

                workSheet.Cells["A" + index + ":B" + index].Merge = true;

                workSheet.Cells["A" + index].Value ="Importante: ";
                workSheet.Cells["A" + index].Style.Font.Size = 20;
                workSheet.Cells["A" + index].Style.Font.Bold = true;

                workSheet.Cells["C" + index].Value = cotizacion.Foot_Importante;
                workSheet.Cells["C" + index].Style.Font.Size = 18;
                workSheet.Cells["C" + index].Style.Font.Bold = true;

                index = index + 2;

                workSheet.Cells["B" + index + ":C" + (index + 6)].Merge = true;
                workSheet.Cells["B" + index + ":C" + (index + 6)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["B" + (index + 7) + ":C" + (index + 8)].Merge = true;
                workSheet.Cells["B" + (index + 7) + ":C" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["B" + (index + 7)].Value= cotizacion.Foot_Firma;
                workSheet.Cells["B" + (index + 7)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B" + (index + 7)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + (index + 7)].Style.Font.Bold = true;
                workSheet.Cells["B" + (index + 7)].Style.Font.Size = 11;

                TextoNegrita(workSheet);

                workSheet.View.ZoomScale = 80;
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
            workSheet.Column(2).Width = 23.71 + 0.71;
            workSheet.Column(3).Width = 43.71 + 0.71;
            workSheet.Column(4).Width = 5.14 + 0.71;
            workSheet.Column(5).Width = 12.57 + 0.71;
            workSheet.Column(6).Width = 9.86 + 0.71;
            workSheet.Column(7).Width = 10.43 + 0.71;
            workSheet.Column(8).Width = 19.29 + 0.71;
            workSheet.Column(9).Width = 9.0 + 0.71;
            workSheet.Column(10).Width = 7.86 + 0.71;
            workSheet.Column(11).Width = 16.29 + 0.71;
            workSheet.Column(12).Width = 12.86 + 0.71;
            workSheet.Column(13).Width = 30.86 + 0.71;
            workSheet.Column(14).Width = 12.57 + 0.71;
            workSheet.Column(15).Width = 21.43 + 0.71;
            workSheet.Column(16).Width = 13.86 + 0.71;
            workSheet.Column(17).Width = 14.14 + 0.71;
            workSheet.Column(18).Width = 18.71 + 0.71;

            workSheet.Column(19).Width = 10.14 + 0.71;
            workSheet.Column(20).Width = 10.14 + 0.71;
            workSheet.Column(21).Width = 12.71 + 0.71;
            workSheet.Column(22).Width = 22 + 0.71;

            workSheet.Row(2).Height = 24.75;
            workSheet.Row(3).Height = 30;
            workSheet.Row(4).Height = 24.75;

            workSheet.Row(6).Height = 5.25;

            workSheet.Row(7).Height = 34.5;

            workSheet.Row(8).Height = 48;
            workSheet.Row(9).Height = 48;
            workSheet.Row(10).Height = 55.5;
            workSheet.Row(11).Height = 38.25;
            workSheet.Row(12).Height = 38.25;
            workSheet.Row(13).Height = 15; 
            workSheet.Row(14).Height = 15; 

            workSheet.Row(15).Height = 48.75; 
            workSheet.Row(16).Height = 28.5; 
            workSheet.Row(17).Height = 60; 
            workSheet.Row(18).Height = 168; 
            workSheet.Row(19).Height = 33.75; 



        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["S2:V2"].Merge = true;
            workSheet.Cells["S3:V3"].Merge = true;
            workSheet.Cells["S4:V4"].Merge = true;
            workSheet.Cells["A7:L7"].Merge = true;
            workSheet.Cells["A8:B8"].Merge = true;
            workSheet.Cells["A9:B9"].Merge = true;
            workSheet.Cells["A10:B10"].Merge = true;
            workSheet.Cells["A11:B11"].Merge = true;
            workSheet.Cells["A12:B12"].Merge = true;

            workSheet.Cells["D8:E8"].Merge = true;
            workSheet.Cells["D9:E9"].Merge = true;
            workSheet.Cells["D10:E10"].Merge = true;
            workSheet.Cells["D11:E11"].Merge = true;
            workSheet.Cells["D12:F12"].Merge = true;
            workSheet.Cells["G12:L12"].Merge = true;

            workSheet.Cells["F8:G8"].Merge = true;
            workSheet.Cells["F9:G9"].Merge = true;
            workSheet.Cells["F10:G10"].Merge = true;
            workSheet.Cells["F11:G11"].Merge = true;

            workSheet.Cells["H8:I8"].Merge = true;
            workSheet.Cells["H9:I9"].Merge = true;
            workSheet.Cells["H10:I10"].Merge = true;
            workSheet.Cells["H11:I11"].Merge = true;

            workSheet.Cells["J8:L8"].Merge = true;
            workSheet.Cells["J9:L9"].Merge = true;
            workSheet.Cells["J10:L10"].Merge = true;
            workSheet.Cells["J11:L11"].Merge = true;

            workSheet.Cells["N7:R7"].Merge = true;
            workSheet.Cells["N8:O8"].Merge = true;
            workSheet.Cells["N9:O9"].Merge = true;
            workSheet.Cells["N10:O10"].Merge = true;
            workSheet.Cells["N11:O11"].Merge = true;
            workSheet.Cells["N12:O12"].Merge = true;

            workSheet.Cells["P8:R8"].Merge = true;
            workSheet.Cells["P9:R9"].Merge = true;
            workSheet.Cells["P10:R10"].Merge = true;
            workSheet.Cells["P11:R11"].Merge = true;
            workSheet.Cells["P12:R12"].Merge = true;


            workSheet.Cells["A15:A19"].Merge = true;
            workSheet.Cells["B15:D16"].Merge = true;
            workSheet.Cells["B17:B19"].Merge = true;
            workSheet.Cells["C17:C19"].Merge = true;
            workSheet.Cells["D17:D19"].Merge = true;

            workSheet.Cells["E15:E19"].Merge = true;

            workSheet.Cells["F15:S15"].Merge = true;
            workSheet.Cells["F16:F19"].Merge = true;
            workSheet.Cells["G16:G19"].Merge = true;
            workSheet.Cells["H16:H19"].Merge = true;
            workSheet.Cells["I16:I19"].Merge = true;
            workSheet.Cells["J16:J19"].Merge = true;
            workSheet.Cells["K16:R16"].Merge = true;
            workSheet.Cells["K17:K19"].Merge = true;
            workSheet.Cells["L17:L19"].Merge = true;
            workSheet.Cells["M17:M19"].Merge = true;
            workSheet.Cells["N17:N19"].Merge = true;
            workSheet.Cells["O17:O19"].Merge = true;
            workSheet.Cells["P17:P19"].Merge = true;
            workSheet.Cells["Q17:Q19"].Merge = true;
            workSheet.Cells["R17:R19"].Merge = true;
            workSheet.Cells["S17:S19"].Merge = true;

            workSheet.Cells["T15:T19"].Merge = true;
            workSheet.Cells["U15:U19"].Merge = true;
            workSheet.Cells["V15:V19"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["C2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A2:B4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D2:R4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S2:V2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S3:V3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S4:V4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A7:L7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A8:B8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A9:B9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A10:B10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A11:B11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A12:B12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D8:E8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D9:E9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D10:E10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D11:E11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F8:G8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F9:G9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F10:G10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F11:G11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D12:L12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H8:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H9:I9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H10:I10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H11:I11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J8:L8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J9:L9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J10:L10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J11:L11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N7:R7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N8:O8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N9:O9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N10:O10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N11:O11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N12:O12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P8:R8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P9:R9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P10:R10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P11:R11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P12:R12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A15:A19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B15:D16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B17:B19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C17:C19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D17:D19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["E15:E19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["F15:S15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F16:F19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G16:G19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H16:H19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I16:I19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J16:J19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K16:R16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K17:K19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L17:L19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M17:M19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N17:N19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O17:O19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P17:P19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q17:Q19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R17:R19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S17:S19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["T15:T19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["U15:U19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["V15:V19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void TextoNegrita(ExcelWorksheet workSheet)
        {
            workSheet.Cells["C2"].Style.Font.Bold = true;
            workSheet.Cells["C3"].Style.Font.Bold = true;
            workSheet.Cells["C4"].Style.Font.Bold = true;
            workSheet.Cells["S2"].Style.Font.Bold = true;
            workSheet.Cells["S3"].Style.Font.Bold = true;
            workSheet.Cells["S4"].Style.Font.Bold = true;
            workSheet.Cells["A7"].Style.Font.Bold = true;
            workSheet.Cells["A8"].Style.Font.Bold = true;
            workSheet.Cells["A9"].Style.Font.Bold = true;
            workSheet.Cells["A10"].Style.Font.Bold = true;
            workSheet.Cells["A11"].Style.Font.Bold = true;
            workSheet.Cells["A12"].Style.Font.Bold = true;
            workSheet.Cells["C12"].Style.Font.Bold = true;
            workSheet.Cells["D8"].Style.Font.Bold = true;
            workSheet.Cells["D9"].Style.Font.Bold = true;
            workSheet.Cells["D10"].Style.Font.Bold = true;
            workSheet.Cells["D11"].Style.Font.Bold = true;
            workSheet.Cells["D12"].Style.Font.Bold = true;
            workSheet.Cells["H8"].Style.Font.Bold = true;
            workSheet.Cells["H9"].Style.Font.Bold = true;
            workSheet.Cells["H10"].Style.Font.Bold = true;
            workSheet.Cells["H11"].Style.Font.Bold = true;
            workSheet.Cells["N7"].Style.Font.Bold = true;
            workSheet.Cells["N8"].Style.Font.Bold = true;
            workSheet.Cells["N9"].Style.Font.Bold = true;
            workSheet.Cells["N10"].Style.Font.Bold = true;
            workSheet.Cells["N11"].Style.Font.Bold = true;
            workSheet.Cells["N12"].Style.Font.Bold = true;

            workSheet.Cells["A15:V19"].Style.Font.Bold = true;

        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["A7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DAEEF3"));

            workSheet.Cells["N7:R7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["N7:R7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DAEEF3"));

            workSheet.Cells["A15:V19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["A15:V19"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DAEEF3"));
        }
    }
}
