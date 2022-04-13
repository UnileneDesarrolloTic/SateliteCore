using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato65_Report
    {
        public static string Exportar(Image logoUnilene, Formato65_Model cotizacion)
        {
            byte[] file;

            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Almenara");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 5, 0, 5);
                imagenUnilene.SetSize(204, 85);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 11;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                BordesCeldas(worksheet);
                TextoNegrita(worksheet);

                worksheet.Cells["J2"].Value = "Fecha:";
                worksheet.Cells["J3"].Value = "Página:";
                worksheet.Cells["K2"].Value = cotizacion.Fecha_1.ToString("dd/MM/yyyy");
                worksheet.Cells["K3"].Value = "1 de 1";

                worksheet.Cells["A4"].Value = "Cotización Nº " + cotizacion.Nro_Cotizacion;
                worksheet.Cells["A4"].Style.Font.Size = 16;

                worksheet.Cells["B5"].Value = "UNILENE S.A.C.";
                worksheet.Cells["B5"].Style.Font.Size = 14;

                worksheet.Cells["B6"].Value = "JR. NAPO 450 - BREÑA";
                worksheet.Cells["B6"].Style.Font.Size = 12;

                worksheet.Cells["B7"].Value = "Lima - Perú";
                worksheet.Cells["B7"].Style.Font.Size = 12;

                worksheet.Cells["A10"].Value = "Código";
                worksheet.Cells["C10"].Value = ":";
                worksheet.Cells["D10"].Value = cotizacion.Codigo;

                worksheet.Cells["A11"].Value = "Señores";
                worksheet.Cells["C11"].Value = ":";
                worksheet.Cells["D11"].Value = cotizacion.Seniores;
                worksheet.Cells["D11"].Style.WrapText = true;

                worksheet.Cells["A12"].Value = "R.U.C.";
                worksheet.Cells["C12"].Value = ":";
                worksheet.Cells["D12"].Value = cotizacion.Ruc;

                worksheet.Cells["A13"].Value = "Dirección";
                worksheet.Cells["C13"].Value = ":";
                worksheet.Cells["D13"].Value = cotizacion.Direccion;
                worksheet.Cells["D13"].Style.WrapText = true;

                worksheet.Cells["A14"].Value = "Teléfono";
                worksheet.Cells["C14"].Value = ":";
                worksheet.Cells["D14"].Value = cotizacion.Telefono;

                worksheet.Cells["A15"].Value = "Fax";
                worksheet.Cells["C15"].Value = ":";
                worksheet.Cells["D15"].Value = cotizacion.Fax;

                worksheet.Cells["A16"].Value = "Atención";
                worksheet.Cells["C16"].Value = ":";
                worksheet.Cells["D16"].Value = cotizacion.Atencion;
                worksheet.Cells["D16"].Style.WrapText = true;

                worksheet.Cells["A17"].Value = "Condición de venta";
                worksheet.Cells["C17"].Value = ":";
                worksheet.Cells["D17"].Value = cotizacion.Condicion;

                worksheet.Cells["A18"].Value = "Fecha";
                worksheet.Cells["C18"].Value = ":";
                worksheet.Cells["D18"].Value = cotizacion.Fecha_2.ToString("dd/MM/yyyy");

                worksheet.Cells["F10"].Value = "Asesor comercial";
                worksheet.Cells["G10"].Value = ":";
                worksheet.Cells["H10"].Value = cotizacion.Asesor;
                worksheet.Cells["H10"].Style.WrapText = true;

                worksheet.Cells["F11"].Value = "E-mail";
                worksheet.Cells["G11"].Value = ":";
                worksheet.Cells["H11"].Value = cotizacion.Email_asesor;
                worksheet.Cells["H11"].Style.WrapText = true;

                worksheet.Cells["F12"].Value = "Teléfono";
                worksheet.Cells["G12"].Value = ":";
                worksheet.Cells["H12"].Value = cotizacion.Telefono_asesor;

                worksheet.Cells["F13"].Value = "Validez de la oferta";
                worksheet.Cells["G13"].Value = ":";
                worksheet.Cells["H13"].Value = cotizacion.Validez_oferta;

                worksheet.Cells["F14"].Value = "Plazo de entrega";
                worksheet.Cells["G14"].Value = ":";
                worksheet.Cells["H14"].Value = cotizacion.Plazo_entrega;

                worksheet.Cells["F15"].Value = "Lugar de entrega";
                worksheet.Cells["G15"].Value = ":";
                worksheet.Cells["H15"].Value = cotizacion.Lugar_entrega;

                worksheet.Cells["F16"].Value = "INCORTERMS";
                worksheet.Cells["G16"].Value = ":";
                worksheet.Cells["H16"].Value = cotizacion.Incorterms;

                worksheet.Cells["A20"].Value = "#";
                worksheet.Cells["A20"].Style.Font.Size = 12;

                worksheet.Cells["B20"].Value = "CODSUT";
                worksheet.Cells["B20"].Style.Font.Size = 12;

                worksheet.Cells["E20"].Value = "DESCRIPCIÓN";
                worksheet.Cells["E20"].Style.Font.Size = 12;

                worksheet.Cells["F20"].Value = "DESCRIPCIÓN UNILENE";
                worksheet.Cells["F20"].Style.Font.Size = 12;

                worksheet.Cells["I20"].Value = "CANTIDAD";
                worksheet.Cells["I20"].Style.Font.Size = 12;

                worksheet.Cells["J20"].Value = "PREC. UNIT";
                worksheet.Cells["J20"].Style.Font.Size = 12;

                worksheet.Cells["K20"].Value = "MONTO TOTAL";
                worksheet.Cells["K20"].Style.WrapText = true;
                worksheet.Cells["K20"].Style.Font.Size = 12;

                int row = 21;

                foreach (Formato65_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Cells["A" + row].Value = row - 20;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["B" + row + ":D" + row].Merge = true;
                    worksheet.Cells["B" + row].Value = item.Codsut;
                    worksheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + row].Style.WrapText = true;

                    worksheet.Cells["E" + row].Value = item.Descripcion;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + row].Style.WrapText = true;

                    worksheet.Cells["F" + row + ":H" + row].Merge = true;
                    worksheet.Cells["F" + row].Value = item.Descripcion_unilene;
                    worksheet.Cells["F" + row + ":H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + row].Style.WrapText = true;

                    worksheet.Cells["I" + row].Value = item.Cantidad;
                    worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + row].Style.WrapText = true;

                    worksheet.Cells["J" + row].Value = item.Precio_unitario;
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0.0000";
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + row].Style.WrapText = true;

                    worksheet.Cells["K" + row].Value = item.Total;
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.0000";
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K" + row].Style.WrapText = true;

                    row++;
                }

                worksheet.Row(row).Height = 11.25;
                worksheet.Row(row + 1).Height = 11.25;

                worksheet.Cells["I" + row + ":J" + (row + 1)].Merge = true;
                worksheet.Cells["I" + row + ":J" + (row + 1)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + row].Value = "MONTO TOTAL S/.";
                worksheet.Cells["I" + row].Style.Font.Bold = true;
                worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["k" + row + ":k" + (row + 1)].Merge = true;
                worksheet.Cells["K" + row + ":k" + (row + 1)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Value = cotizacion.Monto_total;
                worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row += 3;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value= "CONDICIONALES GENERALES";
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "BENEFICIARIO : UNILENE S.A.C.";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "DIRECCIÓN : JR NAPO N° 450 LIMA 05-LIMA";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "BBVA CONTINENTAL";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "DIRECCIÓN : AV. REPUBLIC OF PANAMÁ  3055 - SAN ISIDRO PERÚ";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "MONEDA US $ : N° 0011-0910-0100050517";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "SWIFT CODE BCONPEPL";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Style.Font.Size = 10;

                AlineacionesTexto(worksheet);

                worksheet.View.ZoomScale = 80;
                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }

            return reporte;
        }

        private static void AlineacionesTexto(ExcelWorksheet worksheet)
        {
            worksheet.Cells["K2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["D10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; 
            worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A4,B5,B6,B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A20,B20:D20,E20,F20:H20,I20,J20,K20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["A10:B18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A20,B20:D20,E20,F20:H20,I20,J20,K20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 6.86 + 0.71;
            worksheet.Column(2).Width = 13.86 + 0.71;
            worksheet.Column(3).Width = 2.29 + 0.71;
            worksheet.Column(4).Width = 13.14 + 0.71;
            worksheet.Column(5).Width = 45 + 0.71;
            worksheet.Column(6).Width = 30 + 0.71; 
            worksheet.Column(7).Width = 2.29 + 0.71;
            worksheet.Column(8).Width = 12.86 + 0.71;
            worksheet.Column(9).Width = 12.57 + 0.71;
            worksheet.Column(10).Width = 13.71 + 0.71;
            worksheet.Column(11).Width = 12.86 + 0.71;

            worksheet.Row(10).Height = 21;
            worksheet.Row(11).Height = 21;
            worksheet.Row(12).Height = 21;
            //worksheet.Row(13).Height = 21;
            worksheet.Row(14).Height = 21;
            worksheet.Row(15).Height = 21;
            worksheet.Row(16).Height = 21;
            worksheet.Row(17).Height = 21;
            worksheet.Row(18).Height = 21;
            worksheet.Row(20).Height = 31.5;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A4:K4"].Merge = true;
            worksheet.Cells["A10:B10,A11:B11,A12:B12,A13:B13,A14:B14,A15:B15,A16:B16,A17:B17,A18:B18"].Merge = true;

            worksheet.Cells["D10:E10,D11:E11,D12:E12,D13:E13,D14:E14,D15:E15,D16:E16,D17:E17,D18:E18"].Merge = true;
            worksheet.Cells["H10:K10,H11:K11,H12:K12,H13:K13,H14:K14,H15:K15,H16:K16"].Merge = true;
            worksheet.Cells["B20:D20,F20:H20"].Merge = true;

        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A20,B20:D20,E20,F20:H20,I20,J20,K20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
            worksheet.Cells["J2,J3"].Style.Font.Bold = true;
            worksheet.Cells["A4,B5"].Style.Font.Bold = true;
            worksheet.Cells["A10:C18,F10:G18"].Style.Font.Bold = true;
            worksheet.Cells["A20,B20:D20,E20,F20:H20,I20,J20,K20"].Style.Font.Bold = true;

            worksheet.Cells["A20"].Style.Font.Bold = true;
        }
    }
}
