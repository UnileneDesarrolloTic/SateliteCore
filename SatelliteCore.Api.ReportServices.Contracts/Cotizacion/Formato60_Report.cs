using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public static class Formato60_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Formato60_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Red Prestacional Almenara");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(2, 2, 0, 11);
                imagenUnilene.SetSize(190, 69);

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);

                workSheet.Cells["C2"].Value = "UNIDAD DE PROGRAMACIÓN";
                workSheet.Cells["C2"].Style.Font.Size = 22;
                workSheet.Cells["C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["C2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C3"].Value = "OF. ASTECIMIENTO Y CP";
                workSheet.Cells["C3"].Style.Font.Size = 22;
                workSheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["C3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C4"].Value = "RED PRESTACIONAL ALMENARA";
                workSheet.Cells["C4"].Style.Font.Size = 22;
                workSheet.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["C4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D2"].Value = "SOLICITUD DE COTIZACIÓN";
                workSheet.Cells["D2"].Style.Font.Size = 36;
                workSheet.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["D2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["D2"].Value = "FORMATO DE COTIZACIÓN";
                workSheet.Cells["D2"].Style.Font.Size = 36;
                workSheet.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["M2"].Value = "SOLICITUD DE COTIZACIÓN";
                workSheet.Cells["M2"].Style.Font.Size = 18;

                workSheet.Cells["M3"].Value = "Versión: 01-2020";
                workSheet.Cells["M3"].Style.Font.Size = 18;

                workSheet.Cells["M4"].Value = "Página: 1 / 1";
                workSheet.Cells["M4"].Style.Font.Size = 18;
                workSheet.Cells["M4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A7"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A7"].Style.Font.Size = 28;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A8"].Value = "RAZON SOCIAL";
                workSheet.Cells["A8"].Style.Font.Size = 18;
                workSheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C8"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C8"].Style.Font.Size = 18;
                workSheet.Cells["C8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A9"].Value = "R.U.C.";
                workSheet.Cells["A9"].Style.Font.Size = 18;
                workSheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C9"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["C9"].Style.Font.Size = 18;
                workSheet.Cells["C9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A10"].Value = "DIRECCIÓN";
                workSheet.Cells["A10"].Style.Font.Size = 18;
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C10"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["C10"].Style.Font.Size = 18;
                workSheet.Cells["C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["A11"].Value = "E-MAIL";
                workSheet.Cells["A11"].Style.Font.Size = 18;
                workSheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C11"].Value = cotizacion.Prov_Email;
                workSheet.Cells["C11"].Style.Font.Size = 18;
                workSheet.Cells["C11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C11"].Style.Font.UnderLine = true;
                workSheet.Cells["C11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["A12"].Value = "N° COTIZACIÓN";
                workSheet.Cells["A12"].Style.Font.Size = 18;
                workSheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["C12"].Value = cotizacion.Prov_NroCotizacion;
                workSheet.Cells["C12"].Style.Font.Size = 18;
                workSheet.Cells["C12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C12"].Style.Font.Color.SetColor(Color.Red);

                workSheet.Cells["D8"].Value = "TELÉFONO";
                workSheet.Cells["D8"].Style.Font.Size = 20;
                workSheet.Cells["D8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F8"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["F8"].Style.Font.Size = 18;
                workSheet.Cells["F8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["D9"].Value = "FAX";
                workSheet.Cells["D9"].Style.Font.Size = 20;
                workSheet.Cells["D9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F9"].Value = cotizacion.Prov_Fax;
                workSheet.Cells["F9"].Style.Font.Size = 18;
                workSheet.Cells["F9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["D10"].Value = "TELÉFONO";
                workSheet.Cells["D10"].Style.Font.Size = 20;
                workSheet.Cells["D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F10"].Value = cotizacion.Prov_Telefono2;
                workSheet.Cells["F10"].Style.Font.Size = 18;
                workSheet.Cells["F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["D11"].Value = "FECHA";
                workSheet.Cells["D11"].Style.Font.Size = 20;
                workSheet.Cells["D11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["F11"].Value = cotizacion.Prov_Fecha;
                workSheet.Cells["F11"].Style.Font.Size = 18;
                workSheet.Cells["F11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F11"].Style.Numberformat.Format = "dd/mm/yyyy";

                workSheet.Cells["D12"].Value = "DATOS ADICIONALES";
                workSheet.Cells["D12"].Style.Font.Size = 20;
                workSheet.Cells["D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["G12"].Value = cotizacion.Prov_DatosAdicionales;
                workSheet.Cells["G12"].Style.Font.Size = 18;
                workSheet.Cells["G12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["G8"].Value = "CONTACTO";
                workSheet.Cells["G8"].Style.Font.Size = 20;
                workSheet.Cells["G8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["I8"].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["I8"].Style.Font.Size = 18;
                workSheet.Cells["I8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["G9"].Value = "CARGO";
                workSheet.Cells["G9"].Style.Font.Size = 20;
                workSheet.Cells["G9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["I9"].Value = cotizacion.Prov_Cargo;
                workSheet.Cells["I9"].Style.Font.Size = 18;
                workSheet.Cells["I9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["G10"].Value = "TELÉFONO";
                workSheet.Cells["G10"].Style.Font.Size = 20;
                workSheet.Cells["G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["I10"].Value = cotizacion.Prov_Telefono3;
                workSheet.Cells["I10"].Style.Font.Size = 18;
                workSheet.Cells["I10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["G11"].Value = "E-MAIL";
                workSheet.Cells["G11"].Style.Font.Size = 20;
                workSheet.Cells["G11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["I11"].Value = cotizacion.Prov_Email2;
                workSheet.Cells["I11"].Style.Font.Size = 18;
                workSheet.Cells["I11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I11"].Style.Font.UnderLine = true;
                workSheet.Cells["I11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["K7"].Value = "DATOS DEL ÁREA SOLICITANTE";
                workSheet.Cells["K7"].Style.Font.Size = 20;
                workSheet.Cells["K7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["K8"].Value = "ÁREA USUARIA";
                workSheet.Cells["K8"].Style.Font.Size = 20;
                workSheet.Cells["K8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["M8"].Value = cotizacion.Soli_AreaUsuaria;
                workSheet.Cells["M8"].Style.Font.Size = 18;
                workSheet.Cells["M8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["K9"].Value = "DEPENDENCIA";
                workSheet.Cells["K9"].Style.Font.Size = 20;
                workSheet.Cells["K9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["M9"].Value = cotizacion.Soli_Dependencia;
                workSheet.Cells["M9"].Style.Font.Size = 18;
                workSheet.Cells["M9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["K10"].Value = "CONTACTO";
                workSheet.Cells["K10"].Style.Font.Size = 20;
                workSheet.Cells["K10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["M10"].Value = cotizacion.Soli_Contacto;
                workSheet.Cells["M10"].Style.Font.Size = 18;
                workSheet.Cells["M10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["K11"].Value = "E-MAIL";
                workSheet.Cells["K11"].Style.Font.Size = 20;
                workSheet.Cells["K11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["M11"].Value = cotizacion.Soli_Email;
                workSheet.Cells["M11"].Style.Font.Size = 18;
                workSheet.Cells["M11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M11"].Style.Font.UnderLine = true;
                workSheet.Cells["M11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                workSheet.Cells["K12"].Value = "TELÉFONO";
                workSheet.Cells["K12"].Style.Font.Size = 20;
                workSheet.Cells["K12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["M12"].Value = cotizacion.Soli_Telefono;
                workSheet.Cells["M12"].Style.Font.Size = 18;
                workSheet.Cells["M12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["A15:R16"].Style.Font.Size = 18;

                workSheet.Cells["B15"].Value = "DESCRIPCIÓN DEL ÍTEM";
                workSheet.Cells["B15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["A15"].Value = "N° ITEM";
                workSheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A15"].Style.TextRotation = 90;

                workSheet.Cells["B16"].Value = "CÓDIGO SAP";
                workSheet.Cells["B16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B16"].Style.TextRotation = 90;

                workSheet.Cells["C16"].Value = "DENOMINACIÓN";
                workSheet.Cells["C16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["D16"].Value = "UNIDAD DE MEDIDA";
                workSheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D16"].Style.TextRotation = 90;

                workSheet.Cells["E15"].Value = "Cantidad Solicitada";
                workSheet.Cells["E15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E15"].Style.TextRotation = 90;

                workSheet.Cells["F15"].Value = "REQUERIMIENTOS MÍNIMOS";
                workSheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["F16"].Value = "Marca";
                workSheet.Cells["F16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["G16"].Value = "País de Procedencia";
                workSheet.Cells["G16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G16"].Style.TextRotation = 90;

                workSheet.Cells["H16"].Value = "Presentación del Producto";
                workSheet.Cells["H16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H16"].Style.TextRotation = 90;

                workSheet.Cells["I16"].Value = "Fecha de Vencimiento del Producto (indicar fecha de expiración)";
                workSheet.Cells["I16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I16"].Style.WrapText = true;

                workSheet.Cells["J16"].Value = "Plazo de entrega";
                workSheet.Cells["J16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J16"].Style.TextRotation = 90;

                workSheet.Cells["K16"].Value = "Vigencia de la Oferta";
                workSheet.Cells["K16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K16"].Style.TextRotation = 90;

                workSheet.Cells["L16"].Value = "Forma de Pago";
                workSheet.Cells["L16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L16"].Style.TextRotation = 90;


                workSheet.Cells["M15"].Value = "RNP Vigente a la fecha (SI - NO)";
                workSheet.Cells["M15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M15"].Style.TextRotation = 90;

                workSheet.Cells["N15"].Value = "Periodo de garantía (Mínimo 01 año) ";
                workSheet.Cells["N15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N15"].Style.TextRotation = 90;

                workSheet.Cells["O15"].Value = "Cumple al 100% con las especificaciones técnicas y condiciones generales (Si ó No)";
                workSheet.Cells["O15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O15"].Style.TextRotation = 90;
                workSheet.Cells["O15"].Style.WrapText = true;

                workSheet.Cells["P15"].Value = "COSTO UNITARIO S/. (incluido IGV)";
                workSheet.Cells["P15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P15"].Style.WrapText = true;

                workSheet.Cells["R15"].Value = "VALOR TOTAL A OFERTAR S/. (Incluido IGV)";
                workSheet.Cells["R15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R15"].Style.WrapText = true;

                int row = 17;

                foreach (Formato60_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Cells["A" + row].Value = item.NroItem;
                    workSheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + row].Style.Numberformat.Format = "0";
                    workSheet.Cells["A" + row].Style.Font.Size = 26;
                    workSheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["B" + row].Value = item.CodigoSap;
                    workSheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + row].Style.WrapText = true;
                    workSheet.Cells["B" + row].Style.Font.Size = 18;
                    workSheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + row].Value = item.Denominacion;
                    workSheet.Cells["C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + row].Style.WrapText = true;
                    workSheet.Cells["C" + row].Style.Font.Size = 20;
                    workSheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + row].Value = item.UnidadMedida;
                    workSheet.Cells["D" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + row].Style.TextRotation = 90;
                    workSheet.Cells["D" + row].Style.Font.Size = 18;
                    workSheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + row].Value = item.CantidadSolicitada;
                    workSheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + row].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells["E" + row].Style.Font.Size = 18;
                    workSheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + row].Value = item.Marca;
                    workSheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + row].Style.WrapText = true;
                    workSheet.Cells["F" + row].Style.Font.Size = 18;
                    workSheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + row].Value = item.Procedencia;
                    workSheet.Cells["G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + row].Style.WrapText = true;
                    workSheet.Cells["G" + row].Style.Font.Size = 18;
                    workSheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + row].Value = item.Presentacion;
                    workSheet.Cells["H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + row].Style.WrapText = true;
                    workSheet.Cells["H" + row].Style.Font.Size = 18;
                    workSheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + row].Value = item.Vencimiento;
                    workSheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + row].Style.WrapText = true;
                    workSheet.Cells["I" + row].Style.Font.Size = 18;
                    workSheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + row].Value = item.PlazoEntrega;
                    workSheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + row].Style.WrapText = true;
                    workSheet.Cells["J" + row].Style.Font.Size = 18;
                    workSheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["k" + row].Value = item.Vigencia;
                    workSheet.Cells["k" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["k" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["k" + row].Style.WrapText = true;
                    workSheet.Cells["k" + row].Style.Font.Size = 18;
                    workSheet.Cells["k" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + row].Value = item.FormaPago;
                    workSheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + row].Style.WrapText = true;
                    workSheet.Cells["L" + row].Style.Font.Size = 18;
                    workSheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + row].Value = item.RNP;
                    workSheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + row].Style.WrapText = true;
                    workSheet.Cells["M" + row].Style.Font.Size = 18;
                    workSheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["N" + row].Value = item.PlazoEntrega;
                    workSheet.Cells["N" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + row].Style.WrapText = true;
                    workSheet.Cells["N" + row].Style.Font.Size = 18;
                    workSheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["O" + row].Value = item.EspTecnica;
                    workSheet.Cells["O" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + row].Style.WrapText = true;
                    workSheet.Cells["O" + row].Style.Font.Size = 18;
                    workSheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + row + ":Q" + row].Merge = true;
                    workSheet.Cells["P" + row + ":Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + row].Value = item.CostoUnitario;
                    workSheet.Cells["P" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + row].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["P" + row].Style.Font.Size = 18;

                    workSheet.Cells["R" + row].Value = item.CostoTotal;
                    workSheet.Cells["R" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells["R" + row].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["R" + row].Style.Font.Bold = true;
                    workSheet.Cells["R" + row].Style.Font.Size = 18;
                    workSheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    row++;
                }

                row += 16;
                workSheet.Row(row).Height = 52.5 + 0.71; 
                workSheet.Cells["C" + row].Value = "FIRMA DEL REPRESENTANTE LEGAL DE LA EMPRESA O EL QUE HAGA SUS VECES";
                workSheet.Cells["C" + row].Style.WrapText = true;
                workSheet.Cells["C" + row].Style.Font.Bold = true;
                workSheet.Cells["C" + row].Style.Font.Size = 18;
                workSheet.Cells["C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + row].Style.Border.Top.Style = ExcelBorderStyle.Medium;


                row -= 12;
                ExcelPicture firmaCatizacion = workSheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(row, 0, 2, 140);
                firmaCatizacion.SetSize(360, 175);



                TextoNegrita(workSheet);

                workSheet.View.ZoomScale = 55;

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
            workSheet.Column(2).Width = 23.14 + 0.71;
            workSheet.Column(3).Width = 81.86 + 0.71;
            workSheet.Column(4).Width = 11.14 + 0.71;
            workSheet.Column(5).Width = 15.14 + 0.71;
            workSheet.Column(6).Width = 23.86 + 0.71;
            workSheet.Column(7).Width = 17.14 + 0.71;
            workSheet.Column(8).Width = 20.71 + 0.71;
            workSheet.Column(9).Width = 30.71 + 0.71;
            workSheet.Column(10).Width = 20 + 0.71;
            workSheet.Column(11).Width = 18.71 + 0.71;
            workSheet.Column(12).Width = 18.29 + 0.71;
            workSheet.Column(13).Width = 14.57 + 0.71;
            workSheet.Column(14).Width = 14 + 0.71;
            workSheet.Column(15).Width = 23.29 + 0.71;
            workSheet.Column(16).Width = 10.71 + 0.71;
            workSheet.Column(17).Width = 14.86 + 0.71;
            workSheet.Column(18).Width = 33.71 + 0.71;

            workSheet.Row(1).Height = 6;
            workSheet.Row(2).Height = 24.75;
            workSheet.Row(3).Height = 24.75;
            workSheet.Row(4).Height = 54.75;
            workSheet.Row(5).Height = 15;
            workSheet.Row(6).Height = 15;
            workSheet.Row(7).Height = 38.25;
            workSheet.Row(8).Height = 38.25;
            workSheet.Row(9).Height = 38.25;
            workSheet.Row(10).Height = 52.5;
            workSheet.Row(11).Height = 38.25;
            workSheet.Row(12).Height = 38.25;
            workSheet.Row(13).Height = 15;
            workSheet.Row(14).Height = 15;
            workSheet.Row(15).Height = 36.75;
            workSheet.Row(16).Height = 308;
        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A2:B4, D2:L4, M2:O2, M3:O3, M4:O4"].Merge = true;
            workSheet.Cells["A7:J7"].Merge = true;
            workSheet.Cells["A8:B8, A9:B9, A10:B10, A11:B11, A12:B12"].Merge = true;
            workSheet.Cells["D8:E8, D9:E9, D10:E10, D11:E11, D12:F12"].Merge = true;
            workSheet.Cells["G8:H8, G9:H9, G10:H10, G11:H11, G12:J12"].Merge = true;
            workSheet.Cells["I8:J8, I9:J9, I10:J10, I11:J11"].Merge = true;
            workSheet.Cells["K7:O7, K8:L8, K9:L9, K10:L10, K11:L11, K12:L12"].Merge = true;
            workSheet.Cells["M8:O8, M9:O9, M10:O10, M11:O11, M12:O12"].Merge = true;

            workSheet.Cells["A15:A16, B15:D15, E15:E16, F15:L15, M15:M16, N15:N16, O15:O16, P15:Q16, R15:R16"].Merge = true;


        }
        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A2:B4, D2:L4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C2, C3, C4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M2:O2, M3:O3, M4:O4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A7:J7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A8:B8, A9:B9, A10:B10, A11:B11, A12:B12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D8:E8, D9:E9, D10:E10, D11:E11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C8, C9, C10, C11, C12, F8, F9, F10, F11, D12:F12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G8:H8, G9:H9, G10:H10, G11:H11, G12:J12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I8:J8, I9:J9, I10:J10, I11:J11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K7:O7, K8:L8, K9:L9, K10:L10, K11:L11, K12:L12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M8:O8, M9:O9, M10:O10, M11:O11, M12:O12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A15:A16, B15:D15, E15:E16, F15:L15, M15:M16, N15:N16, O15:O16, P15:Q16, R15:R16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B16, C16, D16, F16, G16, H16, I16, J16, K16, L16, M16, N16, O16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A15:R16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["A15:R16"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#92CDDC"));
        }
        private static void TextoNegrita(ExcelWorksheet workSheet)
        {
            workSheet.Cells["C2, C3, C4, D2, M2, M3, M4"].Style.Font.Bold = true;
            workSheet.Cells["A7, A8, A9, A10, A11, A12, C12"].Style.Font.Bold = true;
            workSheet.Cells["D8, D9, D10, D11, D12"].Style.Font.Bold = true;
            workSheet.Cells["F15, G8, G9, G10, G11"].Style.Font.Bold = true;
            workSheet.Cells["K7, K8, K9, K10, K11, K12, A15:R16"].Style.Font.Bold = true;
        }
    }
}