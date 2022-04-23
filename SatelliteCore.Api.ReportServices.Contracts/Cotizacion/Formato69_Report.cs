using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato69_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Formato69_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Almenara");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 1, 0, 1);
                imagenUnilene.SetSize(87, 40);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 10;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);

                worksheet.Cells["T1"].Value = "Fecha:";
                worksheet.Cells["T1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["U1"].Value = cotizacion.Fecha_documento;
                worksheet.Cells["U1"].Style.Numberformat.Format = "dd/MM/yyyy";
                worksheet.Cells["U1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A3"].Value = $"Cotiz Nº {cotizacion.Nro_cotizacion}-Unilene S.A.C";
                worksheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A4"].Value = "OFICINAS: JR. NAPO 450\nCENTRAL TELEFONICA: 7487006 / 7487000\ncontactenos @unilene.com\nwww.unilene.com contactenosunilene.com";
                worksheet.Cells["A4"].Style.WrapText = true;
                worksheet.Cells["A4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A6"].Value = "Código";
                worksheet.Cells["A6"].Style.Font.Size = 8;
                worksheet.Cells["B6"].Value = ":";
                worksheet.Cells["C6"].Value = cotizacion.Codigo_cliente;
                worksheet.Cells["C6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["C6"].Style.Font.Size = 8;

                worksheet.Cells["A7"].Value = "Señor(es)";
                worksheet.Cells["A7"].Style.Font.Size = 8;
                worksheet.Cells["B7"].Value = ":";
                worksheet.Cells["C7"].Value = cotizacion.Cliente_razonSocial;
                worksheet.Cells["C7"].Style.Font.Size = 8;
                worksheet.Cells["C7"].Style.WrapText = true;

                worksheet.Cells["A8"].Value = "R.U.C.";
                worksheet.Cells["A8"].Style.Font.Size = 8;
                worksheet.Cells["B8"].Value = ":";
                worksheet.Cells["C8"].Value = cotizacion.Ruc_cliente;
                worksheet.Cells["C8"].Style.Font.Size = 8;

                worksheet.Cells["A9"].Value = "Dirección";
                worksheet.Cells["A9"].Style.Font.Size = 8;
                worksheet.Cells["B9"].Value = ":";
                worksheet.Cells["C9"].Value = cotizacion.Direccion_cliente;
                worksheet.Cells["C9"].Style.Font.Size = 8;
                worksheet.Cells["C9"].Style.WrapText = true;

                worksheet.Cells["A10"].Value = "Teléfono";
                worksheet.Cells["A10"].Style.Font.Size = 8;
                worksheet.Cells["B10"].Value = ":";
                worksheet.Cells["C10"].Value = cotizacion.Telefono_cliente;
                worksheet.Cells["C10"].Style.Font.Size = 8;

                worksheet.Cells["A12"].Value = "Fax";
                worksheet.Cells["A12"].Style.Font.Size = 8;
                worksheet.Cells["B12"].Value = ":";
                worksheet.Cells["C12"].Value = cotizacion.Fax_cliente;
                worksheet.Cells["C12"].Style.Font.Size = 8;

                worksheet.Cells["A13"].Value = "Atención";
                worksheet.Cells["A13"].Style.Font.Size = 8;
                worksheet.Cells["B13"].Value = ":";
                worksheet.Cells["C13"].Value = cotizacion.Atencion;
                worksheet.Cells["C13"].Style.Font.Size = 8;

                worksheet.Cells["A14"].Value = "Cond.\nVenta";
                worksheet.Cells["A14"].Style.Font.Size = 8;
                worksheet.Cells["A14"].Style.WrapText = true;
                worksheet.Cells["B14"].Value = ":";
                worksheet.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C14"].Value = cotizacion.Cond_venta;
                worksheet.Cells["C14"].Style.Font.Size = 8;
                worksheet.Cells["C14"].Style.WrapText = true;
                worksheet.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A15"].Value = "Fecha";
                worksheet.Cells["A15"].Style.Font.Size = 8;
                worksheet.Cells["B15"].Value = ":";
                worksheet.Cells["C15"].Value = cotizacion.Fecha;
                worksheet.Cells["C15"].Style.Numberformat.Format = "dd/MM/yyyy";
                worksheet.Cells["C15"].Style.Font.Size = 8;
                worksheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["J6"].Value = "Asesor Comercial";
                worksheet.Cells["J6"].Style.Font.Size = 8;
                worksheet.Cells["O6"].Value = ":";
                worksheet.Cells["P6"].Value = cotizacion.Asesor_prov;
                worksheet.Cells["P6"].Style.Font.Size = 8;

                worksheet.Cells["J7"].Value = "Email";
                worksheet.Cells["J7"].Style.Font.Size = 8;
                worksheet.Cells["O7"].Value = ":";
                worksheet.Cells["P7"].Value = cotizacion.Email_prov;
                worksheet.Cells["P7"].Style.Font.Size = 8;

                worksheet.Cells["J8"].Value = "Teléfono";
                worksheet.Cells["J8"].Style.Font.Size = 8;
                worksheet.Cells["O8"].Value = ":";
                worksheet.Cells["P8"].Value = cotizacion.Telefono_prov;
                worksheet.Cells["P8"].Style.Font.Size = 8;

                worksheet.Cells["J9"].Value = "Validez de la Oferta";
                worksheet.Cells["J9"].Style.Font.Size = 8;
                worksheet.Cells["O9"].Value = ":";
                worksheet.Cells["P9"].Value = cotizacion.Validez_oferta;
                worksheet.Cells["P9"].Style.Font.Size = 8;

                worksheet.Cells["J10"].Value = "Plazo de Entrega";
                worksheet.Cells["J10"].Style.Font.Size = 8;
                worksheet.Cells["O10"].Value = ":";
                worksheet.Cells["P10"].Value = cotizacion.Plazo_entrega;
                worksheet.Cells["P10"].Style.Font.Size = 8;

                worksheet.Cells["J12"].Value = "Lugar de Entrega";
                worksheet.Cells["J12"].Style.Font.Size = 8;
                worksheet.Cells["O12"].Value = ":";
                worksheet.Cells["P12"].Value = cotizacion.Lugar_entrega;
                worksheet.Cells["P12"].Style.Font.Size = 8;

                worksheet.Cells["J13"].Value = "Garantía";
                worksheet.Cells["J13"].Style.Font.Size = 8;
                worksheet.Cells["O13"].Value = ":";
                worksheet.Cells["P13"].Value = cotizacion.Garantia;
                worksheet.Cells["P13"].Style.Font.Size = 8;

                worksheet.Cells["J14"].Value = "Vigencia Producto";
                worksheet.Cells["J14"].Style.Font.Size = 8;
                worksheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O14"].Value = ":";
                worksheet.Cells["P14"].Value = cotizacion.Vigencia_producto;
                worksheet.Cells["P14"].Style.Font.Size = 8;
                worksheet.Cells["P14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["J15"].Value = "Ref";
                worksheet.Cells["J15"].Style.Font.Size = 8;
                worksheet.Cells["O15"].Value = ":";
                worksheet.Cells["P15"].Value = cotizacion.Ref;
                worksheet.Cells["P15"].Style.Font.Size = 8;

                worksheet.Cells["A17"].Value = "COD. SUT";
                worksheet.Cells["A17"].Style.Font.Size = 8;
                worksheet.Cells["A17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A17"].Style.HorizontalAlignment= ExcelHorizontalAlignment.Center;

                worksheet.Cells["A17"].Value = "COD. SUT";
                worksheet.Cells["A17"].Style.Font.Size = 8;
                worksheet.Cells["A17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B17"].Value = "DESCRIPCIÓN";
                worksheet.Cells["B17"].Style.Font.Size = 8;
                worksheet.Cells["B17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F17"].Value = "MARCA";
                worksheet.Cells["F17"].Style.Font.Size = 8;
                worksheet.Cells["F17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["I17"].Value = "PRESENTACIÓN";
                worksheet.Cells["I17"].Style.Font.Size = 8;
                worksheet.Cells["I17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["M17"].Value = "PROCED.";
                worksheet.Cells["M17"].Style.Font.Size = 8;
                worksheet.Cells["M17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["N17"].Value = "UND.\nMED";
                worksheet.Cells["N17"].Style.WrapText = true;
                worksheet.Cells["N17"].Style.Font.Size = 8;
                worksheet.Cells["N17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["Q17"].Value = "CANT.\nUNITARIA";
                worksheet.Cells["Q17"].Style.WrapText = true;
                worksheet.Cells["Q17"].Style.Font.Size = 8;
                worksheet.Cells["Q17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["R17"].Value = "PLAZO\nENTREGA";
                worksheet.Cells["R17"].Style.WrapText = true;
                worksheet.Cells["R17"].Style.Font.Size = 8;
                worksheet.Cells["R17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["S17"].Value = "PRECIO\nUNITARIO";
                worksheet.Cells["S17"].Style.WrapText = true;
                worksheet.Cells["S17"].Style.Font.Size = 8;
                worksheet.Cells["S17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["V17"].Value = "TOTAL";
                worksheet.Cells["V17"].Style.WrapText = true;
                worksheet.Cells["V17"].Style.Font.Size = 8;
                worksheet.Cells["V17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["V17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 18;

                foreach (Formato69_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Cells["A" + row].Value = item.Cod_sut;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["B" + row + ":E" + row].Merge = true;
                    worksheet.Cells["B" + row].Value = item.Descripcion;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + row + ":E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["F" + row + ":G" + row].Merge = true;
                    worksheet.Cells["F" + row].Value = item.Marca;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + row + ":G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["I" + row + ":L" + row].Merge = true;
                    worksheet.Cells["I" + row].Value = item.Presentacion;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + row + ":L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["M" + row].Value = item.Procendencia;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["N" + row + ":P" + row].Merge = true;
                    worksheet.Cells["N" + row].Value = item.Unidad_medida;
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    worksheet.Cells["N" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["N" + row + ":P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["Q" + row].Value = item.Cantidad;
                    worksheet.Cells["Q" + row].Style.WrapText = true;
                    worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["Q" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["R" + row].Value = item.Plazo_entrega;
                    worksheet.Cells["R" + row].Style.WrapText = true;
                    worksheet.Cells["R" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["S" + row + ":U" + row].Merge = true;
                    worksheet.Cells["S" + row].Value = item.Precio_unitario;
                    worksheet.Cells["S" + row].Style.WrapText = true;
                    worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["S" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["S" + row + ":U" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["V" + row + ":W" + row].Merge = true;
                    worksheet.Cells["V" + row].Value = item.Total;
                    worksheet.Cells["V" + row].Style.WrapText = true;
                    worksheet.Cells["V" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["V" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["V" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["V" + row + ":W" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    row++;
                }

                worksheet.Cells["Q" + row + ":R" + row].Merge = true;
                worksheet.Cells["Q" + row + ":R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["Q" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q" + row].Style.Font.Size = 8;
                worksheet.Cells["Q" + row].Value = "Monto Total S/.";
                worksheet.Cells["Q" + row].Style.Font.Bold = true;

                worksheet.Cells["S" + row + ":W" + row].Merge = true;
                worksheet.Cells["S" + row + ":W" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["S" + row].Value = cotizacion.Monto_total;
                worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["S" + row].Style.Font.Bold = true;

                row++;

                worksheet.Row(row).Height = 18;
                worksheet.Cells["A" + row + ":J" + row].Merge = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                worksheet.Cells["A" + row].Value = "CONDICIONES GENERALES:";
                worksheet.Cells["A" + row].Style.Font.UnderLine = true;
                worksheet.Cells["A" + row].Style.Font.Bold = true;

                row++;

                worksheet.Row(row).Height = 17;
                worksheet.Cells["A" + row + ":D" + row].Merge = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Value = "PRECIO";
                worksheet.Cells["A" + row].Style.Font.Size = 8;

                worksheet.Cells["E" + row + ":J" + row].Merge = true;
                worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E" + row].Value = "EN Soles INCLUIDO I.G.V";
                worksheet.Cells["E" + row].Style.Font.Size = 8;

                row++;

                worksheet.Row(row).Height = 17;
                worksheet.Cells["A" + row + ":D" + row].Merge = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Value = "FORMA DE PAGO";
                worksheet.Cells["A" + row].Style.Font.Size = 8;

                worksheet.Cells["E" + row + ":J" + row].Merge = true;
                worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E" + row].Value = "Crédito 90 días";
                worksheet.Cells["E" + row].Style.Font.Size = 8;

                row++;

                worksheet.Row(row).Height = 17;
                worksheet.Cells["A" + row + ":D" + row].Merge = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Value = "LABORATORIO FABRICANTE";
                worksheet.Cells["A" + row].Style.Font.Size = 8;

                worksheet.Cells["E" + row + ":J" + row].Merge = true;
                worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E" + row].Value = "UNILENE S.A.C.";
                worksheet.Cells["E" + row].Style.Font.Size = 8;

                row++;

                worksheet.Row(row).Height = 17;
                worksheet.Cells["A" + row + ":D" + row].Merge = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Value = "R.U.C.";
                worksheet.Cells["A" + row].Style.Font.Size = 8;

                worksheet.Cells["E" + row + ":J" + row].Merge = true;
                worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["E" + row].Value = "20197705249";
                worksheet.Cells["E" + row].Style.Font.Size = 8;

                row++;

                worksheet.Row(row).Height = 17;
                worksheet.Cells["A" + row + ":D" + row].Merge = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + row].Value = "FECHA DE VENCIMIENTO DE COTIZACION";
                worksheet.Cells["A" + row].Style.Font.Size = 8;

                worksheet.Cells["E" + row + ":J" + row].Merge = true;
                worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["E" + row].Value = cotizacion.Fecha_vencimiento;
                worksheet.Cells["E" + row].Style.Numberformat.Format = "dd/MM/yyyy";
                worksheet.Cells["E" + row].Style.Font.Size = 8;

                row += 2;

                worksheet.Row(row).Height = 24.75;
                worksheet.Cells["A" + row + ":M" + row].Merge = true;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Value = "Sin otro particular y a la espera de sus gratas ordenes, quedamos de Ustedes.\nAtentamente,";
                worksheet.Cells["A" + row].Style.Font.Size = 8;

                row ++;

                ExcelPicture firmaExcel = worksheet.Drawings.AddPicture("firma", firma);
                firmaExcel.SetPosition(row, 0, 0, 15);
                firmaExcel.SetSize(190, 90);

                row += 7;

                worksheet.Row(row).Height = 37;
                worksheet.Cells["A" + row + ":W" + row].Merge = true;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Value = "NOTA : Vencida la cotización comunicarse con el Area Comercial de nuestra empresa a los teléfonos señalados a fin de actualizar la información vía E-mail indicando la nueva fecha de entrega de nuestros productos y así poder atenderlos evitando retrasos innecesarios UNILENE NO SE HACE RESPONSABLE SI VENCIDA LA COTIZACIÓN REMITEN ORDENES DE COMPRA SIN COORDINACION PREVIA DE LA ENTREGA";

                TextoNegrita(worksheet);
                BordesCeldas(worksheet);

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
            worksheet.Column(1).Width = 9.71 + 0.71;
            worksheet.Column(2).Width = 0.5 + 0.71;
            worksheet.Column(3).Width = 8.71 + 0.71;
            worksheet.Column(4).Width = 11 + 0.71;
            worksheet.Column(5).Width = 6 + 0.71;
            worksheet.Column(6).Width = 8.57 + 0.71;
            worksheet.Column(7).Width = 2.57 + 0.71;
            worksheet.Column(8).Width = 0;
            worksheet.Column(9).Width = 6.29 + 0.71;
            worksheet.Column(10).Width = 1.14 + 0.71;
            worksheet.Column(11).Width = 0;
            worksheet.Column(12).Width = 2.71 + 0.71;
            worksheet.Column(13).Width = 9 + 0.71;
            worksheet.Column(14).Width = 1.43 + 0.71;
            worksheet.Column(15).Width = 0.5 + 0.71;
            worksheet.Column(16).Width = 1.57 + 0.71;
            worksheet.Column(17).Width = 9.57 + 0.71;
            worksheet.Column(18).Width = 9.43 + 0.71;
            worksheet.Column(19).Width = 0.75 + 0.71;
            worksheet.Column(20).Width = 5.71 + 0.71;
            worksheet.Column(21).Width = 2.29 + 0.71;
            worksheet.Column(22).Width = 6 + 0.71;
            worksheet.Column(23).Width = 3.86 + 0.71;

            worksheet.Row(2).Height = 6.25;
            worksheet.Row(3).Height = 17;
            worksheet.Row(4).Height = 47.75;
            worksheet.Row(5).Height = 5;

            worksheet.Row(11).Height = 0;
            worksheet.Row(14).Height = 22.5;
            worksheet.Row(17).Height = 29;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:W3"].Merge = true;
            worksheet.Cells["A4:G4,U1:W1"].Merge = true;
            worksheet.Cells["C6:G6,C7:G7,C8:G8,C9:G9,C10:G10,C12:G12,C13:G13,C14:G14,C15:G15"].Merge = true;
            worksheet.Cells["J6:N6,J7:N7,J8:N8,J9:N9,J10:N10,J12:N12,J13:N13,J14:N14,J15:N15"].Merge = true;
            worksheet.Cells["P6:V6,P7:V7,P8:V8,P9:V9,P10:V10,P12:V12,P13:V13,P14:V14,P15:V15"].Merge = true;
            worksheet.Cells["B17:E17,F17:G17,I17:L17,N17:P17,R17,Q17,S17:U17,V17:W17"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B17:E17,F17:G17,I17:L17,M17,N17:P17,Q17,R17:U17,V17:W17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
            worksheet.Cells["U1,T1,A3"].Style.Font.Bold = true;
            worksheet.Cells["A6:B15,J6:O15,A17:W17"].Style.Font.Bold = true;
        }
    }
}
