using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato68_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Coti_Formato_68_Model cotizacion)
        {
            byte[] file;

            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Mariana del Peru");

                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                imagenUnilene.SetPosition(3, 0, 1, 4);
                imagenUnilene.SetSize(450, 250);

                workSheet.Cells.Style.Font.Name = "Century Gothic";
                workSheet.Cells.Style.Font.Size = 30;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);


                workSheet.Cells["I7"].Value = "COTIZ .Nro " + cotizacion.Prov_NroCotizacion +"-2022 Unilene S.A.C";
                workSheet.Cells["I7"].Style.Font.Size = 30;
                workSheet.Cells["I7"].Style.Font.UnderLine = true;
                workSheet.Cells["I7"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["I7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["I7"].Style.Font.Bold = true;

                workSheet.Cells["B9"].Value ="INDAGACIÓN DE PRECIOS EN EL MERCADO - DISPOSITIVO MÉDICO";
                workSheet.Cells["B9"].Style.Font.Size =50;
                workSheet.Cells["B9"].Style.Font.UnderLine = true;
                workSheet.Cells["B9"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B9"].Style.Font.Bold = true;

                workSheet.Cells["B11"].Value = "ATENCIÓN : " + cotizacion.Prov_atencion;
                workSheet.Cells["B11"].Style.Font.Size = 30;
                workSheet.Cells["B11"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B11"].Style.Font.Bold = true;


                workSheet.Cells["B15"].Value = "FECHA DE ELABORACIÓN :";
                workSheet.Cells["B15"].Style.Font.Size = 30;
                workSheet.Cells["B15"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B15"].Style.Font.Bold = true;

                workSheet.Cells["D15"].Value = cotizacion.Prov_Fecha.ToLongDateString();
                workSheet.Cells["D15"].Style.Font.Size = 30;
                workSheet.Cells["D15"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D15"].Style.Font.Bold = true;

                workSheet.Cells["B16"].Value = "RAZÓN SOCIAL :";
                workSheet.Cells["B16"].Style.Font.Size = 30;
                workSheet.Cells["B16"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B16"].Style.Font.Bold = true;

                workSheet.Cells["D16"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["D16"].Style.Font.Size = 30;
                workSheet.Cells["D16"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D16"].Style.Font.Bold = true;

                workSheet.Cells["B17"].Value = "RUC :";
                workSheet.Cells["B17"].Style.Font.Size = 30;
                workSheet.Cells["B17"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B17"].Style.Font.Bold = true;

                workSheet.Cells["D17"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["D17"].Style.Font.Size = 30;
                workSheet.Cells["D17"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D17"].Style.Font.Bold = true;


                workSheet.Cells["B18"].Value = "CORREO ELECTRÓNICO:";
                workSheet.Cells["B18"].Style.Font.Size = 30;
                workSheet.Cells["B18"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B18"].Style.Font.Bold = true;

                workSheet.Cells["D18"].Value = cotizacion.Prov_Email;
                workSheet.Cells["D18"].Style.Font.Size = 30;
                workSheet.Cells["D18"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D18"].Style.Font.Bold = true;

                workSheet.Cells["B19"].Value = "CONTACTO:";
                workSheet.Cells["B19"].Style.Font.Size = 30;
                workSheet.Cells["B19"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B19"].Style.Font.Bold = true;

                workSheet.Cells["D19"].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["D19"].Style.Font.Size = 30;
                workSheet.Cells["D19"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D19"].Style.Font.Bold = true;

                workSheet.Cells["B20"].Value = "TELÉFONO MOVIL:";
                workSheet.Cells["B20"].Style.Font.Size = 30;
                workSheet.Cells["B20"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B20"].Style.Font.Bold = true;

                workSheet.Cells["D20"].Value = cotizacion.Prov_telefono;
                workSheet.Cells["D20"].Style.Font.Size = 30;
                workSheet.Cells["D20"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D20"].Style.Font.Bold = true;

                workSheet.Cells["B21"].Value = "TELÉFONO MOVIL:";
                workSheet.Cells["B21"].Style.Font.Size = 30;
                workSheet.Cells["B21"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B21"].Style.Font.Bold = true;

                workSheet.Cells["D21"].Value = cotizacion.Prov_telefono;
                workSheet.Cells["D21"].Style.Font.Size = 30;
                workSheet.Cells["D21"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D21"].Style.Font.Bold = true;


                workSheet.Cells["B22"].Value = "REPRESENTANTE DE VENTAS:";
                workSheet.Cells["B22"].Style.Font.Size = 30;
                workSheet.Cells["B22"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B22"].Style.Font.Bold = true;

                workSheet.Cells["D22"].Value = cotizacion.Prov_RepVentas;
                workSheet.Cells["D22"].Style.Font.Size = 30;
                workSheet.Cells["D22"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D22"].Style.Font.Bold = true;


                workSheet.Cells["B23"].Value = "NOMBRE DEL BANCO DE SU REPRESENTADA:";
                workSheet.Cells["B23"].Style.Font.Size = 30;
                workSheet.Cells["B23"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B23"].Style.Font.Bold = true;

                workSheet.Cells["D23"].Value = cotizacion.Prov_NombreBanco;
                workSheet.Cells["D23"].Style.Font.Size = 30;
                workSheet.Cells["D23"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D23"].Style.Font.Bold = true;


                workSheet.Cells["B24"].Value = "CCI:";
                workSheet.Cells["B24"].Style.Font.Size = 30;
                workSheet.Cells["B24"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B24"].Style.Font.Bold = true;

                workSheet.Cells["D24"].Value = cotizacion.Prov_cci;
                workSheet.Cells["D24"].Style.Font.Size = 30;
                workSheet.Cells["D24"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D24"].Style.Font.Bold = true;


                workSheet.Cells["B25"].Value = "REPRESENTANTE DE VENTAS:";
                workSheet.Cells["B25"].Style.Font.Size = 30;
                workSheet.Cells["B25"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B25"].Style.Font.Bold = true;

                workSheet.Cells["D25"].Value = cotizacion.Prov_EntidadBancaria;
                workSheet.Cells["D25"].Style.Font.Size = 30;
                workSheet.Cells["D25"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D25"].Style.Font.Bold = true;

                workSheet.Cells["B26"].Value = "PERIODO DE VALIDEZ DE SU OFERTA:";
                workSheet.Cells["B26"].Style.Font.Size = 30;
                workSheet.Cells["B26"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B26"].Style.Font.Bold = true;

                workSheet.Cells["D26"].Value = cotizacion.Prov_valiOferta;
                workSheet.Cells["D26"].Style.Font.Size = 30;
                workSheet.Cells["D26"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["D26"].Style.Font.Bold = true;


                workSheet.Cells["B28"].Value = "REFERENCIA DE LA SOLICITUD: "+ cotizacion.Prov_referencia;
                workSheet.Cells["B28"].Style.Font.Size = 50;
                workSheet.Cells["B28"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B28"].Style.Font.Bold = true;

                //DETALLE


                workSheet.Cells["B30"].Value = "ITEM";
                workSheet.Cells["B30"].Style.Font.Size = 30;
                workSheet.Cells["B30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["B30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B30"].Style.WrapText = true;

                workSheet.Cells["C30"].Value = "DESCRIPCION DEL PRODUCTO A COTIZAR";
                workSheet.Cells["C30"].Style.Font.Size = 30;
                workSheet.Cells["C30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C30"].Style.WrapText = true;

                workSheet.Cells["D30"].Value = "UNIDAD DE MEDIDA";
                workSheet.Cells["D30"].Style.Font.Size = 30;
                workSheet.Cells["D30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["D30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D30"].Style.WrapText = true;

                workSheet.Cells["E30"].Value = "CANTIDAD";
                workSheet.Cells["E30"].Style.Font.Size = 30;
                workSheet.Cells["E30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["E30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E30"].Style.WrapText = true;

                workSheet.Cells["F30"].Value = "INDICAR NOMBRE COMERCIAL SEGÚN SU DESCRIPCION EN FACTURA Y GUIA";
                workSheet.Cells["F30"].Style.Font.Size = 30;
                workSheet.Cells["F30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["F30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F30"].Style.WrapText = true;

                workSheet.Cells["G30"].Value = "MARCA";
                workSheet.Cells["G30"].Style.Font.Size = 30;
                workSheet.Cells["G30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["G30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G30"].Style.WrapText = true;

                workSheet.Cells["H30"].Value = "PROCED.";
                workSheet.Cells["H30"].Style.Font.Size = 30;
                workSheet.Cells["H30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["H30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H30"].Style.WrapText = true;

                workSheet.Cells["I30"].Value = "PRESENTA";
                workSheet.Cells["I30"].Style.Font.Size = 30;
                workSheet.Cells["I30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["I30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I30"].Style.WrapText = true;

                workSheet.Cells["J30"].Value = "PRECIO UNITARIO INCLUIDO IGV  S/.";
                workSheet.Cells["J30"].Style.Font.Size = 30;
                workSheet.Cells["J30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["J30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J30"].Style.WrapText = true;

                workSheet.Cells["K30"].Value = "PRECIO TOTAL INCLUIDO IGV  S/.";
                workSheet.Cells["K30"].Style.Font.Size = 30;
                workSheet.Cells["K30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["K30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K30"].Style.WrapText = true;

                workSheet.Cells["L30"].Value = "VIGENCIA DEL PRODUCTO (MIN. 12 MESES) Ó PRESENTAR CARTA DE CANJE SEGÚN SEA EL CASO";
                workSheet.Cells["L30"].Style.Font.Size = 30;
                workSheet.Cells["L30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["L30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L30"].Style.WrapText = true;


                workSheet.Cells["M30"].Value = "CUMPLE CON LAS FICHAS TÉCNICAS ó ESPECIFICACIONES TÉCNICAS     (INDICAR SI Ó NO)";
                workSheet.Cells["M30"].Style.Font.Size = 30;
                workSheet.Cells["M30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["M30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M30"].Style.WrapText = true;

                workSheet.Cells["N30"].Value = "GARANTIA DEL BIEN (MÍNIMO 12 MESES)";
                workSheet.Cells["N30"].Style.Font.Size = 30;
                workSheet.Cells["N30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["N30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N30"].Style.WrapText = true;


                workSheet.Cells["O30"].Value = "PLAZO DE ENTREGA DEL BIEN";
                workSheet.Cells["O30"].Style.Font.Size = 30;
                workSheet.Cells["O30"].Style.Font.Name = "Century Gothic";
                workSheet.Cells["O30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O30"].Style.WrapText = true;


                int index = 31;

                foreach (Coti_Formato68_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Row(index).Height = 150;
                    workSheet.Cells["B" + index].Value = index - 30;
                    workSheet.Cells["B" + index].Style.Font.Size = 26;
                    workSheet.Cells["B" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["B" + index].Style.WrapText = true;


                    workSheet.Cells["C" + index].Value = item.Descripcion;
                    workSheet.Cells["C" + index].Style.Font.Size = 26;
                    workSheet.Cells["C" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["C" + index].Style.WrapText = true;

                    workSheet.Cells["D" + index].Value = item.UndMedida;
                    workSheet.Cells["D" + index].Style.Font.Size = 26;
                    workSheet.Cells["D" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["D" + index].Style.WrapText = true;

                    workSheet.Cells["E" + index].Value = item.Cantidad;
                    workSheet.Cells["E" + index].Style.Font.Size = 26;
                    workSheet.Cells["E" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells["E" + index].Style.WrapText = true;

                    workSheet.Cells["F" + index].Value = item.Nombcomercial;
                    workSheet.Cells["F" + index].Style.Font.Size = 26;
                    workSheet.Cells["F" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;

                    workSheet.Cells["G" + index].Value = item.Marca;
                    workSheet.Cells["G" + index].Style.Font.Size = 26;
                    workSheet.Cells["G" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;


                    workSheet.Cells["H" + index].Value = item.Procedencia;
                    workSheet.Cells["H" + index].Style.Font.Size = 26;
                    workSheet.Cells["H" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;

                    workSheet.Cells["I" + index].Value = item.Presentacion;
                    workSheet.Cells["I" + index].Style.Font.Size = 26;
                    workSheet.Cells["I" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Cells["J" + index].Value = item.PreUnitario;
                    workSheet.Cells["J" + index].Style.Font.Size = 26;
                    workSheet.Cells["J" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;
                    workSheet.Cells["J" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["K" + index].Value = item.PreTotal;
                    workSheet.Cells["K" + index].Style.Font.Size = 26;
                    workSheet.Cells["K" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;
                    workSheet.Cells["K" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["L" + index].Value = item.VigProducto;
                    workSheet.Cells["L" + index].Style.Font.Size = 26;
                    workSheet.Cells["L" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;


                    workSheet.Cells["M" + index].Value = item.CumpFichaTecnica;
                    workSheet.Cells["M" + index].Style.Font.Size = 26;
                    workSheet.Cells["M" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    workSheet.Cells["N" + index].Value = item.Garantia;
                    workSheet.Cells["N" + index].Style.Font.Size = 26;
                    workSheet.Cells["N" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;

                    workSheet.Cells["O" + index].Value = item.Plazoentrega;
                    workSheet.Cells["O" + index].Style.Font.Size = 26;
                    workSheet.Cells["O" + index].Style.Font.Name = "Century Gothic";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;


                    index++;
                }
               
                workSheet.Row(index).Height = 77.25;
                workSheet.Cells["J" + index].Value = "TOTAL S/";
                workSheet.Cells["J" + index].Style.Font.Size = 36;
                workSheet.Cells["J" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["J" + index].Style.WrapText = true;
                

                workSheet.Cells["K" + index].Value = cotizacion.Prov_valorTotal;
                workSheet.Cells["K" + index].Style.Font.Size = 36;
                workSheet.Cells["K" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["K" + index].Style.WrapText = true;
                workSheet.Cells["K" + index].Style.Numberformat.Format = "#,##0.00";

                index++;
                workSheet.Row(index).Height = 38.25;
                workSheet.Cells["C" + index].Value = "NOTA IMPORTANTE";
                workSheet.Cells["C" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index].Style.Font.UnderLine = true;
                workSheet.Cells["C" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index].Style.WrapText = true;


                index++;
                workSheet.Row(index).Height = 120.5;
                workSheet.Cells["C" + index + ":I" + index].Merge = true;
                workSheet.Cells["C" + index + ":I" + index].Value = "Asimismo, declaro no encontrarme impedido para postular en el proceso de contratación, selección ni contratar con el estado , conforme al articulo 11 de la ley N° 30225, ley de contrataciones del estado, asi como conozco las sanciones" +
                    " contenidas en dicha Ley, su Reglamento y la Ley N° 27444, Ley del Procedimiento Administrativo General.";
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["C" + index + ":I" + index].Style.WrapText = true;


                index++;
                workSheet.Row(index).Height = 38.25;
                workSheet.Cells["C" + index + ":I" + index].Merge = true;
                workSheet.Cells["C" + index + ":I" + index].Value = "Es obligatorio anexar las Declaraciones adjuntas debidamente llenadas, (SÓLO PARA LA EMPRESA ADJUDICADA)";
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["C" + index + ":I" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 38.25;
                workSheet.Cells["C" + index + ":I" + index].Merge = true;
                workSheet.Cells["C" + index + ":I" + index].Value = "Nota:Para la formalización contractual u Orden de Compra el proveedor deberá contar con la siguiente documentación:";
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":I" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                workSheet.Cells["C" + index + ":I" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 38.25;

                index++;
                workSheet.Row(index).Height = 65.25;
                workSheet.Cells["C" + index + ":D" + index].Merge = true;
                workSheet.Cells["C" + index + ":D" + index].Value = "DOCUMENTACIÓN";
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E" + index].Merge = true;
                workSheet.Cells["E" + index].Value = "SI Ó  NO";
                workSheet.Cells["E" + index].Style.Font.Size = 30;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Row(index).Height = 38.25;
                workSheet.Cells["C" + index + ":D" + index].Merge = true;
                workSheet.Cells["C" + index + ":D" + index].Value = "Consulta RUC (Verificar CIIU, Estado Activo, Condición habido)";
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                workSheet.Cells["E" + index].Value = "SI";
                workSheet.Cells["E" + index].Style.Font.Size = 30;
                workSheet.Cells["E" + index].Style.Font.Bold = true;
                workSheet.Cells["E" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Row(index).Height = 38.25;
                workSheet.Cells["C" + index + ":D" + index].Merge = true;
                workSheet.Cells["C" + index + ":D" + index].Value = "Consulta RUC (Verificar CIIU, Estado Activo, Condición habido)";
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                workSheet.Cells["E" + index].Value = "SI";
                workSheet.Cells["E" + index].Style.Font.Size = 30;
                workSheet.Cells["E" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Row(index).Height = 38.25;
                workSheet.Cells["C" + index + ":D" + index].Merge = true;
                workSheet.Cells["C" + index + ":D" + index].Value = "Código de Cuenta Interbancaria (CCI)";
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                workSheet.Cells["E" + index].Value = "SI";
                workSheet.Cells["E" + index].Style.Font.Size = 30;
                workSheet.Cells["E" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Row(index).Height = 267.75;
                workSheet.Cells["C" + index + ":D" + index].Merge = true;
                workSheet.Cells["C" + index + ":D" + index].Value = "No tener impedimento para contratar con el Estado";
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Size = 30;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["C" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["C" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                workSheet.Cells["E" + index].Value = "No tebnemos Impedimento para contratar con el Estado";
                workSheet.Cells["E" + index].Style.Font.Size = 30;
                workSheet.Cells["E" + index].Style.Font.Name = "Century Gothic";
                workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + index].Style.WrapText = true;
                workSheet.Cells["E" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
                workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Row(index).Height = 38.25;
                index++;
                workSheet.Row(index).Height = 38.25;

                workSheet.Cells["G" + index + ":K" + (index + 1)].Merge = true;
                workSheet.Cells["G" + index + ":K"+ (index+1)].Value = "FIRMA y SELLO";
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.Font.Size = 30;
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.Font.Bold = true;
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.Font.Name = "Century Gothic";
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.WrapText = true;
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
                workSheet.Cells["G" + index + ":K" + (index + 1)].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#F40000"));


                ExcelPicture firmaCatizacion = workSheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(index-6, -8, 7, 85);
                firmaCatizacion.SetSize(1000, 550);

                workSheet.View.ZoomScale = 25; //80
                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

            }

            return reporte;
        }


        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Column(1).Width = 10.71 + 0.71;
            workSheet.Column(2).Width = 26.00 + 0.71;
            workSheet.Column(3).Width = 152.14 + 0.71;
            workSheet.Column(4).Width = 31.43 + 0.71;
            workSheet.Column(5).Width = 39.43 + 0.71;
            workSheet.Column(6).Width = 102.86 + 0.71; 
            workSheet.Column(7).Width = 40 + 0.71;

            workSheet.Column(8).Width = 40.00 + 0.71;
            workSheet.Column(9).Width = 40.00 + 0.71;
            workSheet.Column(10).Width = 40.00 + 0.71;

            workSheet.Column(11).Width = 40.00 + 0.71;
            workSheet.Column(12).Width = 53.71 + 0.71;
            workSheet.Column(13).Width = 40.00 + 0.71;
            workSheet.Column(14).Width = 46.57 + 0.71;
            workSheet.Column(15).Width = 39.86 + 0.71;

            workSheet.Row(1).Height = 38.25;
            workSheet.Row(2).Height = 38.25;
            workSheet.Row(3).Height = 38.25;
            workSheet.Row(4).Height = 38.25;

            workSheet.Row(5).Height = 36.75;
            workSheet.Row(6).Height = 36.75;
            workSheet.Row(7).Height = 37.50;
            workSheet.Row(8).Height = 36.75;


            workSheet.Row(9).Height = 61.5;
            workSheet.Row(10).Height = 36.75;
            workSheet.Row(11).Height = 36.75;
            workSheet.Row(12).Height = 36.75;


            workSheet.Row(13).Height = 61.5;
            workSheet.Row(14).Height = 36.75;
            workSheet.Row(15).Height = 36.75;
            workSheet.Row(16).Height = 36.75;
            workSheet.Row(17).Height = 36.75;
            workSheet.Row(18).Height = 36.75;
            workSheet.Row(19).Height = 36.75;
            workSheet.Row(20).Height = 36.75;
            workSheet.Row(21).Height = 36.75;
            workSheet.Row(22).Height = 36.75;
            workSheet.Row(23).Height = 36.75;
            workSheet.Row(24).Height = 36.75;
            workSheet.Row(25).Height = 36.75;


            workSheet.Row(26).Height = 36.75;
            workSheet.Row(27).Height = 36.75;
            workSheet.Row(28).Height = 61.5;
            workSheet.Row(29).Height = 39;

            workSheet.Row(30).Height = 373.5;

            


        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["I7:L7"].Merge = true;
            workSheet.Cells["B9:H9"].Merge = true;

            workSheet.Cells["B11:C11"].Merge = true;
            workSheet.Cells["B15:C15"].Merge = true;
            workSheet.Cells["D15:J15"].Merge = true;
            workSheet.Cells["B16:C16"].Merge = true;
            workSheet.Cells["D16:J16"].Merge = true;
            workSheet.Cells["B17:C17"].Merge = true;
            workSheet.Cells["D17:J17"].Merge = true;

            workSheet.Cells["B18:C18"].Merge = true;
            workSheet.Cells["D18:J18"].Merge = true;

            workSheet.Cells["B19:C19"].Merge = true;
            workSheet.Cells["D19:J19"].Merge = true;
            workSheet.Cells["B20:C20"].Merge = true;
            workSheet.Cells["D20:J20"].Merge = true;

            workSheet.Cells["B21:C21"].Merge = true;
            workSheet.Cells["D21:J21"].Merge = true;

            workSheet.Cells["B22:C22"].Merge = true;
            workSheet.Cells["D22:J22"].Merge = true;

            workSheet.Cells["B28:L28"].Merge = true;
        }
        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["B30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            //workSheet.Cells["L4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#333399"));

            workSheet.Cells["B9:H9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
            workSheet.Cells["B28:L28"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
            workSheet.Cells["B30:O30"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
    }
}
