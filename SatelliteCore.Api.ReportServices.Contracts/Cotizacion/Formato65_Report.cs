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
        public static string Exportar(Image firma, Image logoUnilene, Formato65_Model cotizacion)
        {
            byte[] file;

            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Formato General");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 5, 0, 5);
                imagenUnilene.SetSize(204, 85);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 12;
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

                worksheet.Cells["A4"].Value = "Cotización Nº " + cotizacion.Nro_Cotizacion +"-" +cotizacion.Fecha_1.ToString("yyyy") ;
                worksheet.Cells["A4"].Style.Font.Size = 16;

                worksheet.Cells["A5"].Value = "OFICINAS: JR. NAPO 450";
                worksheet.Cells["A5"].Style.Font.Size = 12;
                worksheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A6"].Value = "CENTRAL TELEFÓNICA: 997509088";
                worksheet.Cells["A6"].Style.Font.Size = 12;
                worksheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A7"].Value = "info@unilene.com";
                worksheet.Cells["A7"].Style.Font.Size = 12;
                worksheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A8"].Value = "www.unilene.com";
                worksheet.Cells["A8"].Style.Font.Size = 12;
                worksheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A10"].Value = "Código";
                worksheet.Cells["C10"].Value = ":";
                worksheet.Cells["D10"].Value = cotizacion.Codigo;

                worksheet.Cells["A11"].Value = "Señores";
                worksheet.Cells["C11"].Value = ":";
                worksheet.Cells["D11"].Value = cotizacion.Seniores;
                worksheet.Cells["D11"].Style.WrapText = true;
                worksheet.Cells["D11"].Style.Font.Size = 10;

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
                worksheet.Cells["D18"].Value = cotizacion.Fecha_1.ToString("dd/MM/yyyy");

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

                worksheet.Cells["F16"].Value = "Garantia";
                worksheet.Cells["G16"].Value = ":";
                worksheet.Cells["H16"].Value = cotizacion.Garantia;

                worksheet.Cells["F17"].Value = "Vigencia Producto";
                worksheet.Cells["G17"].Value = ":";
                worksheet.Cells["H17"].Value = cotizacion.Prov_VigProducto;

                worksheet.Cells["F18"].Value = "Ref";
                worksheet.Cells["G18"].Value = ":";
                worksheet.Cells["H18"].Value = cotizacion.Prov_Ref;

          

                //DETALLE
                worksheet.Cells["A20"].Value = "CODSUT";
                worksheet.Cells["A20"].Style.Font.Size = 12;

                worksheet.Cells["D20"].Value = "DESCRIPCIÓN";
                worksheet.Cells["D20"].Style.Font.Size = 12;

                worksheet.Cells["E20"].Value = "MARCA";
                worksheet.Cells["E20"].Style.Font.Size = 12;

                worksheet.Cells["F20"].Value = "PRESENTACION";
                worksheet.Cells["F20"].Style.Font.Size = 12;

                worksheet.Cells["G20"].Value = "PROCED.";
                worksheet.Cells["G20"].Style.Font.Size = 12;

                worksheet.Cells["I20"].Value = "UND MEDIDA";
                worksheet.Cells["I20"].Style.Font.Size = 12;
                worksheet.Cells["I20"].Style.WrapText = true;


                worksheet.Cells["J20"].Value = "CANTIDAD";
                worksheet.Cells["J20"].Style.Font.Size = 12;

                worksheet.Cells["K20"].Value = "PLAZOENTREGA";
                worksheet.Cells["K20"].Style.Font.Size = 11;
                worksheet.Cells["K20"].Style.WrapText = true;

                worksheet.Cells["L20"].Value = "PRECIO UNITARIO";
                worksheet.Cells["L20"].Style.Font.Size = 12;
                worksheet.Cells["L20"].Style.WrapText = true;

                worksheet.Cells["M20"].Value = "TOTAL";
                worksheet.Cells["M20"].Style.Font.Size = 12;
                worksheet.Cells["M20"].Style.WrapText = true;

                //                worksheet.Cells["J20"].Value = "PREC. UNIT";
                //worksheet.Cells["J20"].Style.Font.Size = 12;

                //worksheet.Cells["K20"].Value = "MONTO TOTAL";
                //worksheet.Cells["K20"].Style.WrapText = true;
                // worksheet.Cells["K20"].Style.Font.Size = 12;

                int row = 21;

                foreach (Formato65_Detalle item in cotizacion.Detalle)
                {

                    worksheet.Row(row).Height = 39.75;

                    worksheet.Cells["A" + row + ":C" + row].Merge = true;
                    worksheet.Cells["A" + row + ":C" + row].Value = item.Codsut;
                    worksheet.Cells["A" + row + ":C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + row + ":C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + row + ":C" + row].Style.WrapText = true;

                    worksheet.Cells["D" + row].Value = item.Descripcion;
                    worksheet.Cells["D" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.Font.Size = 12;


                    worksheet.Cells["E" + row].Merge = true;
                    worksheet.Cells["E" + row].Value = item.Marca;
                    worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E" + row].Style.WrapText = true;


                    worksheet.Cells["F" + row].Value = item.Presentacion;
                    worksheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["G" + row + ":H" + row].Merge = true;
                    worksheet.Cells["G" + row + ":H" + row ].Value = item.Procedencia;
                    worksheet.Cells["G" + row + ":H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + row + ":H" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row + ":H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["I" + row].Value = item.Unidadmedida;
                    worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["J" + row].Value = item.Cantidad;
                    worksheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0";

                    worksheet.Cells["K" + row].Value = item.PlazoEntrega;
                    worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + row].Style.Font.Size = 11;


                    worksheet.Cells["L" + row].Value = item.Precio_unitario;
                    worksheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0.00";


                    worksheet.Cells["M" + row].Value = item.Total;
                    worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.00";



                    row++;
                }

                worksheet.Row(row).Height = 11.25;
                worksheet.Row(row + 1).Height = 11.25;

                worksheet.Cells["K" + row + ":L" + (row + 1)].Merge = true;
                worksheet.Cells["K" + row + ":L" + (row + 1)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Value = "MONTO TOTAL S/.";
                worksheet.Cells["K" + row].Style.Font.Bold = true;
                worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["M" + row + ":M" + (row + 1)].Merge = true;
                worksheet.Cells["M" + row + ":M" + (row + 1)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + row].Value = cotizacion.Monto_total;
                worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row += 3;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value= "CONDICIONALES GENERALES";
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "PRECIO : EN Soles INCLUIDO I.G.V.";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "FORMA DE PAGO : " + cotizacion.Condicion;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "LABORATORIO FABRICANTE: UNILENE S.A.C";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "RUC  : 20197705249";
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "FECHA DE VENCIMIENTO DE COTIZACION: " + cotizacion.Fecha_2.ToString("dd/MM/yyyy"); ;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Style.Font.Size = 10;

                row++;
                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Sin otro particular y a la espera de sus gratas ordenes, quedamos de Ustedes:";
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row].Style.Font.Size = 10;

                row++;
                worksheet.Row(row).Height = 19;
                worksheet.Cells["A" + row + ":E" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Atentamente,";
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row].Style.Font.Size = 10;


                row += 7;

                worksheet.Row(row).Height = 25;
                worksheet.Cells["A" + row + ":J" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "NOTA : Vencida la cotización comunicarse con el Area Comercial de nuestra empresa a los teléfonos señalados a fin de actualizar la información vía E-mail indicando"
                    + "la nueva fecha de entrega de nuestros productos y así poder atenderlos evitando retrasos innecesarios";
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row].Style.Font.Size = 10;
                worksheet.Cells["A" + row + ":J" + row].Style.WrapText = true;

                row++;
                worksheet.Cells["A" + row + ":J" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "UNILENE NO SE HACE RESPONSABLE SI VENCIDA LA COTIZACIÓN REMITEN ORDENES DE COMPRA SIN COORDINACION PREVIA DE LA ENTREGA";
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row].Style.Font.Size = 10;


                ExcelPicture firmaCatizacion = worksheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(row-8, -8, 1, 85);
                firmaCatizacion.SetSize(265, 120);

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
            worksheet.Column(4).Width = 34.00 + 0.71;
            worksheet.Column(5).Width = 45 + 0.71;
            worksheet.Column(6).Width = 30 + 0.71; 
            worksheet.Column(7).Width = 2.29 + 0.71;
            worksheet.Column(8).Width = 12.86 + 0.71;
            worksheet.Column(9).Width = 12.57 + 0.71;
            worksheet.Column(10).Width = 13.71 + 0.71;
            worksheet.Column(11).Width = 15.86 + 0.71;
            worksheet.Column(12).Width = 14.86 + 0.71;
            worksheet.Column(13).Width = 14.86 + 0.71;

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
            worksheet.Cells["A5:D5"].Merge = true;
            worksheet.Cells["A6:D6"].Merge = true;
            worksheet.Cells["A7:D7"].Merge = true;
            worksheet.Cells["A8:D8"].Merge = true;

            worksheet.Cells["A10:B10,A11:B11,A12:B12,A13:B13,A14:B14,A15:B15,A16:B16,A17:B17,A18:B18"].Merge = true;

            worksheet.Cells["D10:E10,D11:E11,D12:E12,D13:E13,D14:E14,D15:E15,D16:E16,D17:E17,D18:E18"].Merge = true;
            worksheet.Cells["H10:K10,H11:K11,H12:K12,H13:K13,H14:K14,H15:K15,H16:K16,H17:K17,H18:K18"].Merge = true;
           
            worksheet.Cells["A20:C20"].Merge = true;
            worksheet.Cells["G20:H20"].Merge = true;
            
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A20,B20:D20,E20,F20:H20,I20,J20,K20,L20,M20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
            worksheet.Cells["J2,J3"].Style.Font.Bold = true;
            worksheet.Cells["A4,B5"].Style.Font.Bold = true;
            worksheet.Cells["A10:C18,F10:G18"].Style.Font.Bold = true;
            worksheet.Cells["A20,B20:D20,E20,F20:H20,I20,J20,K20,L20,M20"].Style.Font.Bold = true;

            worksheet.Cells["A20"].Style.Font.Bold = true;
        }
    }
}
