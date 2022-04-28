using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public static class Formato10_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Formato10_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Sabogal");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 50);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                PintarCeldas(worksheet);
                BordesCeldas(worksheet);
                BordesCeldas(worksheet);
                TextoNegrita(worksheet);
                worksheet.Cells["J5"].Value = "Cotiz Nº " + cotizacion.Nro_cotizacion + "-Unilene S.A.C";
                worksheet.Cells["J5"].Style.Font.Size = 10;

                worksheet.Cells["A8"].Value = "SOLICITUD DE COTIZACION - ADQUISICION DE MATERIAL MÉDICO - ESSALUD AREQUIPA";
                worksheet.Cells["A8"].Style.Font.Size = 11;

                worksheet.Cells["A9"].Value = "FECHA: " + cotizacion.Fecha.ToString("dd-MM-yyyy");
                worksheet.Cells["A9"].Style.Font.Size = 10;

                worksheet.Cells["A11"].Value = "DATOS DEL PROVEEDOR";
                worksheet.Cells["A11"].Style.Font.Size = 10;

                worksheet.Cells["A12"].Value = "RAZÓN SOCIAL";
                worksheet.Cells["A12"].Style.Font.Size = 10;

                worksheet.Cells["D12"].Value = cotizacion.Prov_razonSocial;
                worksheet.Cells["D12"].Style.Font.Size = 10;

                worksheet.Cells["A13"].Value = "DIRECCIÓN";
                worksheet.Cells["A13"].Style.Font.Size = 10;

                worksheet.Cells["D13"].Value = cotizacion.Prov_direccion;
                worksheet.Cells["D13"].Style.Font.Size = 10;

                worksheet.Cells["A14"].Value = "TELÉFONO (S)";
                worksheet.Cells["A14"].Style.Font.Size = 10;

                worksheet.Cells["D14"].Value = cotizacion.Prov_telefono;
                worksheet.Cells["D14"].Style.Font.Size = 10;

                worksheet.Cells["A15"].Value = "EMAIL";
                worksheet.Cells["A15"].Style.Font.Size = 10;

                worksheet.Cells["D15"].Value = cotizacion.Prov_email;
                worksheet.Cells["D15"].Style.Font.Size = 10;

                worksheet.Cells["A16"].Value = "VIGENCIA DE LA OFERTA";
                worksheet.Cells["A16"].Style.Font.Size = 10;

                worksheet.Cells["D16"].Value = cotizacion.Vigencia_oferta;
                worksheet.Cells["D16"].Style.Font.Size = 10;

                worksheet.Cells["G16"].Value = "R.U.C.";
                worksheet.Cells["G16"].Style.Font.Size = 10;

                worksheet.Cells["H16"].Value = cotizacion.Prov_ruc;
                worksheet.Cells["H16"].Style.Font.Size = 10;

                worksheet.Cells["M11"].Value = "Contacto";
                worksheet.Cells["M11"].Style.Font.Name = "Calibri";
                worksheet.Cells["M11"].Style.Font.Size = 10;

                worksheet.Cells["N11"].Value = cotizacion.Cont_nombre;
                worksheet.Cells["N11"].Style.Font.Name = "Calibri";
                worksheet.Cells["N11"].Style.Font.Size = 10;

                worksheet.Cells["M12"].Value = "Cargo";
                worksheet.Cells["M12"].Style.Font.Name = "Calibri";
                worksheet.Cells["M12"].Style.Font.Size = 11;

                worksheet.Cells["N12"].Value = cotizacion.Cont_cargo;
                worksheet.Cells["N12"].Style.Font.Name = "Calibri";
                worksheet.Cells["N12"].Style.Font.Size = 11;

                worksheet.Cells["M13"].Value = "Tf.";
                worksheet.Cells["M13"].Style.Font.Name = "Calibri";
                worksheet.Cells["M13"].Style.Font.Size = 11;

                worksheet.Cells["N13"].Value = cotizacion.Cont_telefono;
                worksheet.Cells["N13"].Style.Font.Name = "Calibri";
                worksheet.Cells["N13"].Style.Font.Size = 11;

                worksheet.Cells["M14"].Value = "Mail";
                worksheet.Cells["M14"].Style.Font.Name = "Calibri";
                worksheet.Cells["M14"].Style.Font.Size = 11;

                worksheet.Cells["N14"].Value = cotizacion.Cont_email;
                worksheet.Cells["N14"].Style.Font.Name = "Calibri";
                worksheet.Cells["N14"].Style.Font.Size = 9;

                worksheet.Cells["A19"].Value = "N° \nITEM";
                worksheet.Cells["A19"].Style.Font.Size = 9;
                worksheet.Cells["A19"].Style.WrapText = true;

                worksheet.Cells["B19"].Value = "DESCRIPCIÓN DEL ITEM";
                worksheet.Cells["B19"].Style.Font.Size = 10;

                worksheet.Cells["B20"].Value = "CÓDIGO\nSAP";
                worksheet.Cells["B20"].Style.Font.Size = 9;

                worksheet.Cells["C20"].Value = "DENOMINACIÓN";
                worksheet.Cells["C20"].Style.Font.Size = 9;

                worksheet.Cells["D20"].Value = "UM";
                worksheet.Cells["D20"].Style.Font.Size = 9;

                worksheet.Cells["E20"].Value = "REQUERIM.\nTOTAL";
                worksheet.Cells["E20"].Style.Font.Size = 9;
                worksheet.Cells["E20"].Style.WrapText = true;

                worksheet.Cells["F20"].Value = "PROVEEDOR\n(RAZÓN\nSOCIAL)";
                worksheet.Cells["F20"].Style.Font.Size = 9;
                worksheet.Cells["F20"].Style.WrapText = true;

                worksheet.Cells["G19"].Value = "REQUERIMIENTOS MÍNIMOS";
                worksheet.Cells["G19"].Style.Font.Size = 10;

                worksheet.Cells["G20"].Value = "PLAZO DE ENTREGA EN DÍAS CALENDARIOS";
                worksheet.Cells["G20"].Style.Font.Size = 9;
                worksheet.Cells["G20"].Style.Font.Name = "Tahoma";
                worksheet.Cells["G20"].Style.Font.Bold = false;
                worksheet.Cells["G20"].Style.TextRotation = 90;
                worksheet.Cells["G20"].Style.WrapText = true;

                worksheet.Cells["H20"].Value = "MARCA/ MODELO/ N° PARTE";
                worksheet.Cells["H20"].Style.Font.Size = 9;
                worksheet.Cells["H20"].Style.Font.Name = "Tahoma";
                worksheet.Cells["H20"].Style.Font.Bold = false;
                worksheet.Cells["H20"].Style.TextRotation = 90;
                worksheet.Cells["H20"].Style.WrapText = true;

                worksheet.Cells["I20"].Value = "PROCEDENCIA";
                worksheet.Cells["I20"].Style.Font.Size = 9;
                worksheet.Cells["I20"].Style.Font.Name = "Tahoma";
                worksheet.Cells["I20"].Style.Font.Bold = false;
                worksheet.Cells["I20"].Style.TextRotation = 90;
                worksheet.Cells["I20"].Style.WrapText = true;

                worksheet.Cells["J20"].Value = "CUENTA CON REGISTRO NACIONAL DE PROVEEDORES\n(SI-NO)";
                worksheet.Cells["J20"].Style.Font.Size = 9;
                worksheet.Cells["J20"].Style.Font.Name = "Tahoma";
                worksheet.Cells["J20"].Style.Font.Bold = false;
                worksheet.Cells["J20"].Style.TextRotation = 90;
                worksheet.Cells["J20"].Style.WrapText = true;

                worksheet.Cells["K20"].Value = "CDECLARA CUMPLIR AL 100% CON LOS TDR (SI-NO)";
                worksheet.Cells["K20"].Style.Font.Size = 9;
                worksheet.Cells["K20"].Style.Font.Name = "Tahoma";
                worksheet.Cells["K20"].Style.Font.Bold = false;
                worksheet.Cells["K20"].Style.TextRotation = 90;
                worksheet.Cells["K20"].Style.WrapText = true;

                worksheet.Cells["L20"].Value = "GARANTIA EN MESES";
                worksheet.Cells["L20"].Style.Font.Size = 9;
                worksheet.Cells["L20"].Style.Font.Name = "Tahoma";
                worksheet.Cells["L20"].Style.Font.Bold = false;
                worksheet.Cells["L20"].Style.TextRotation = 90;
                worksheet.Cells["L20"].Style.WrapText = true;

                worksheet.Cells["M19"].Value = "VALOR\nUNITARIO\nEN SOLES\nS/.";
                worksheet.Cells["M19"].Style.Font.Size = 10;
                worksheet.Cells["M19"].Style.WrapText = true;

                worksheet.Cells["N19"].Value = "VALOR TOTAL EN\nSOLES\nS/.";
                worksheet.Cells["N19"].Style.Font.Size = 10;
                worksheet.Cells["N19"].Style.WrapText = true;

                int row = 23;

                foreach (Formato10_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Cells["A" + row].Value = row - 22;
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 12;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["B" + row].Value = item.Sap;
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 11;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["C" + row].Value = item.Denominacion;
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 11;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["D" + row].Value = item.Um;
                    worksheet.Cells["D" + row].Style.Font.Name = "Tahoma";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["E" + row].Value = item.Requerimiento_total;
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 11;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["F" + row].Value = item.Proveedor;
                    worksheet.Cells["F" + row].Style.Font.Name = "Tahoma";
                    worksheet.Cells["F" + row].Style.Font.Size = 11;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["G" + row].Value = item.Plazo_entrega;
                    worksheet.Cells["G" + row].Style.Font.Name = "Tahoma";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["H" + row].Value = item.Marca_modelo;
                    worksheet.Cells["H" + row].Style.Font.Name = "Futura Lt BT";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["I" + row].Value = item.Procedencia;
                    worksheet.Cells["I" + row].Style.Font.Name = "Futura Lt BT";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["J" + row].Value = item.Registro_proveedores;
                    worksheet.Cells["J" + row].Style.Font.Name = "Futura Lt BT";
                    worksheet.Cells["J" + row].Style.Font.Size = 12;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["K" + row].Value = item.Cumple_tdr;
                    worksheet.Cells["K" + row].Style.Font.Name = "Futura Lt BT";
                    worksheet.Cells["K" + row].Style.Font.Size = 12;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["L" + row].Value = item.Garantia_meses;
                    worksheet.Cells["L" + row].Style.Font.Name = "Futura Lt BT";
                    worksheet.Cells["L" + row].Style.Font.Size = 9;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L" + row].Style.TextRotation = 90;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["M" + row].Value = item.ValorUnitario;
                    worksheet.Cells["M" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["M" + row].Style.Font.Size = 10;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.000";
                    worksheet.Cells["M" + row].Style.Font.Bold = true;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["N" + row].Value = item.ValorTotal;
                    worksheet.Cells["N" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["N" + row].Style.Font.Size = 10;
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells["N" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["N" + row].Style.Font.Bold = true;
                    worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    row++;
                }

                worksheet.Cells["M" + row].Value = "TOTAL S/.";
                worksheet.Cells["M" + row].Style.Font.Name = "Arial";
                worksheet.Cells["M" + row].Style.Font.Size = 10;
                worksheet.Cells["M" + row].Style.WrapText = true;
                worksheet.Cells["M" + row].Style.Font.Bold = true;
                worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Medium);


                worksheet.Cells["N" + row].Value = cotizacion.Monto_total;
                worksheet.Cells["N" + row].Style.Font.Name = "Arial";
                worksheet.Cells["N" + row].Style.Font.Size = 10;
                worksheet.Cells["N" + row].Style.WrapText = true;
                worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["N" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["N" + row].Style.Font.Bold = true;
                worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;

                worksheet.Cells["A" + row + ":L" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "1.- El precio de Mercado será a todo costo, es decir, deberá incluir todos los tributos (incluido el I.G.V.), seguros, transportes, inspecciones.";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 39.75;
                worksheet.Cells["A" + row + ":L" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "2.- En el caso que la empresa postora sea a la vez la empresa fabricante nacional; en mérito a la aplicación de los dispositivos que en esta materia se encuentran vigentes en el\nterritorio peruano, de tener Certificado de Buenas Prácticas de Manufactura (CBPM) indicar \"si\", en la columna de Certificado de Buenas Prácticas de Almacenamiento (CBPA),\nentiendiéndose que el CBPM contiene al CBPA. ";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 10.5;
                worksheet.Cells["A" + row + ":L" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "3.- (*) Documentos Alternativos del CBPM: Ó Certificado de Libre Venta ó Certificado de Libre Comercialización.";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;  

                worksheet.Row(row).Height = 12;
                worksheet.Cells["A" + row + ":L" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "5.-De ser el caso; indicar si su Precio de Mercado está exento del I.G.V. y aranceles.";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 36.75;
                worksheet.Cells["A" + row + ":L" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "6.- Excepcionalmente, para los productos que por sus propiedades biológicas, físicas y químicas no puedan cumplir con la vigencia mínima establecida, deben INDICAR en el\nrespectivo recuadro, la palabra \"EXCEPCION\", de ser el caso vigencias menores, siempre que éstas no sean inferiores al 80%  del tiempo de vida útil especificado para el producto\ny declarado por el fabricante.";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 12.75;
                worksheet.Cells["A" + row + ":D" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Lugar de Entrega:";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.Font.Bold = true;
                worksheet.Cells["A" + row].Style.Font.UnderLine = true;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;

                worksheet.Row(row).Height = 22.5;
                worksheet.Cells["A" + row + ":L" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "RED ASISTENCIAL AREQUIPA - (ALMACEN - SITO EN ESQUINA CALLE PERAL CON AYACUCHO).";
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelPicture firmaCatizacion = worksheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(row, -8, 1, 85);
                firmaCatizacion.SetSize(265, 120);

                AlineacionesTexto(worksheet);

                worksheet.View.ZoomScale = 100;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);
            }

            return reporte;
        }

        private static void AlineacionesTexto(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["A11,A12,A13,A14,A15,A16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["A11,D12,D13,D14,D15,D16,G16,H16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A11,D12,D13,D14,D15,D16,G16,H16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["M11,N11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["M11,N11,M12,N12,M13,N13,M14,N14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["A19:N21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A19:N21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 5.43 + 0.71;
            worksheet.Column(2).Width = 9.57 + 0.71;
            worksheet.Column(3).Width = 21.71 + 0.71;
            worksheet.Column(4).Width = 3.86 + 0.71;
            worksheet.Column(5).Width = 9.14 + 0.71;
            worksheet.Column(6).Width = 11.29 + 0.71;
            worksheet.Column(7).Width = 9.71 + 0.71;
            worksheet.Column(8).Width = 16.71 + 0.71;
            worksheet.Column(9).Width = 8.14 + 0.71;
            worksheet.Column(10).Width = 10.29 + 0.71;
            worksheet.Column(11).Width = 7.86 + 0.71;
            worksheet.Column(12).Width = 3.43 + 0.71;
            worksheet.Column(13).Width = 10.29 + 0.71;
            worksheet.Column(14).Width = 20.14 + 0.71;

            worksheet.Row(7).Height = 6;
            worksheet.Row(10).Height = 0;
            worksheet.Row(17).Height = 0;
            worksheet.Row(18).Height = 0;
            worksheet.Row(19).Height = 22.5;
            worksheet.Row(20).Height = 44.25;
            worksheet.Row(21).Height = 42.75;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A8:N8"].Merge = true;
            worksheet.Cells["A11:K11"].Merge = true;
            worksheet.Cells["A12:C12,A13:C13,A14:C14,A15:C15,A16:C16"].Merge = true;
            worksheet.Cells["D12:K12,D13:K13,D14:K14,D15:K15,D16:F16,H16:K16"].Merge = true;

            worksheet.Cells["A19:A22,B19:F19,B20:B22,C20:C22,D20:D22,E20:E22,F20:F22"].Merge = true;
            worksheet.Cells["G19:L19,G20:G22,H20:H22,I20:I22,J20:J22,K20:K22,L20:L22"].Merge = true;
            worksheet.Cells["M19:M22,N19:N22"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A11:K11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A12:C12,A13:C13,A14:C14,A15:C15,A16:C16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["D12:K12,D13:K13,D14:K14,D15:K15,D16:F16,H16:K16,G16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["M12,M13,M14,N12,N13,N14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A11:K16"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells["M11,N11"].Style.Border.BorderAround(ExcelBorderStyle.Medium);

            worksheet.Cells["A19:A22,B19:F19,B20:B22,C20:C22,D20:D22,E20:E22,F20:F22"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells["G19:L19,G20:G22,H20:H22,I20:I22,J20:J22,K20:K22,L20:L22"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells["M19:M22,N19:N22"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["N14,J5"].Style.Font.UnderLine = true;
            worksheet.Cells["D15,N14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

            worksheet.Cells["A19:N21"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A19:N21"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));

        }
        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
            worksheet.Cells["J5,A8,A9"].Style.Font.Bold = true;
            worksheet.Cells["A11,G16"].Style.Font.Bold = true;
            worksheet.Cells["M11,N11"].Style.Font.Bold = true;
            worksheet.Cells["A19:N21"].Style.Font.Bold = true;
        }
    }
}
