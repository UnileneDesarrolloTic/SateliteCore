using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato4_Report
    {
     
        public static string Exportar(Image logoUnilene, Coti_Formato4_Model cotizacion)
        {
                byte[] file;

                string reporte = null;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Almenara");

                    ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                    imagenUnilene.SetPosition(2, 0, 10, 10);
                    imagenUnilene.SetSize(190, 95);


                    workSheet.Cells.Style.Font.Name = "Arial";
                    workSheet.Cells.Style.Font.Size = 10;
                    workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                    workSheet.Cells["B2"].Value = "SOLICITUD DE COTIZACION ";
                    workSheet.Cells["B2"].Style.Font.Bold = true;
                    workSheet.Cells["B2"].Style.Font.Size = 12;

                    workSheet.Cells["N2"].Value = "Fecha Emision";
                    workSheet.Cells["N2"].Style.WrapText = true;
                    workSheet.Cells["N2"].Style.Font.Bold = true;
                    workSheet.Cells["N2"].Style.Font.Size = 10;
                    workSheet.Cells["N2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#16365C"));

                    workSheet.Cells["N4"].Value = cotizacion.Prov_Fecha;
                    workSheet.Cells["N4"].Style.Numberformat.Format = "dd/MM/yyyy";
                    workSheet.Cells["N4"].Style.WrapText = true;
                    workSheet.Cells["N4"].Style.Font.Bold = true;
                    workSheet.Cells["N4"].Style.Font.Size = 10;
                    workSheet.Cells["N4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    workSheet.Cells["B3"].Value = "ÁREA DE PROGRAMACIÓN - OFICINA DE ADQUISICIONES";
                    workSheet.Cells["B3"].Style.Font.Bold = true;
                    workSheet.Cells["B3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B3"].Style.Font.Size = 12;

                    workSheet.Cells["B4"].Value = "RED ASISTENCIAL REBAGLIATI - ESSALUD";
                    workSheet.Cells["B4"].Style.Font.Bold = true;
                    workSheet.Cells["B4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B4"].Style.Font.Size = 24;

                    workSheet.Cells["B6"].Value = "SR. PROVEEDOR:";
                    workSheet.Cells["B6"].Style.Font.Bold = true;
                    workSheet.Cells["B6"].Style.Font.Size = 16;


                    workSheet.Cells["B7"].Value = "Por medio de la presente solicitamos se sirva llenar la solicitud de cotización para precios referenciales de lo siguiente: (en concordancia al Artículo 13° del Reglamento de la Ley de Contrataciones del Estado) ";
                    workSheet.Cells["B7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B7"].Style.Font.Size = 11;

                    workSheet.Cells["B8"].Value = "RAZON SOCIAL:";
                    workSheet.Cells["B8"].Style.Font.Bold = true;
                    workSheet.Cells["B8"].Style.Font.Size = 11;
                    workSheet.Cells["B8"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                    workSheet.Cells["C8"].Value = cotizacion.Prov_RazonSocial;
                    workSheet.Cells["C8"].Style.Font.Bold = true;
                    workSheet.Cells["C8"].Style.Font.Size = 11;
                  

                    workSheet.Cells["B9"].Value = "NUMERO DE RUC:";
                    workSheet.Cells["B9"].Style.Font.Bold = true;
                    workSheet.Cells["B9"].Style.Font.Size = 11;
                    workSheet.Cells["B9"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                    workSheet.Cells["C9"].Value = cotizacion.Prov_Ruc;
                    workSheet.Cells["C9"].Style.Font.Bold = true;
                    workSheet.Cells["C9"].Style.Font.Size = 11;


                    workSheet.Cells["B10"].Value = "DIRECCION:";
                    workSheet.Cells["B10"].Style.Font.Bold = true;
                    workSheet.Cells["B10"].Style.Font.Size = 11;
                    workSheet.Cells["B10"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                    workSheet.Cells["C10"].Value = cotizacion.Prov_Direccion;
                    workSheet.Cells["C10"].Style.Font.Bold = true;
                    workSheet.Cells["C10"].Style.Font.Size = 11;


                    workSheet.Cells["B11"].Value = "TELEFONO:";
                    workSheet.Cells["B11"].Style.Font.Bold = true;
                    workSheet.Cells["B11"].Style.Font.Size = 11;
                    workSheet.Cells["B11"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                    workSheet.Cells["C11"].Value = cotizacion.Prov_Telefono;
                    workSheet.Cells["C11"].Style.Font.Bold = true;
                    workSheet.Cells["C11"].Style.Font.Size = 11;


                    workSheet.Cells["B12"].Value = "CONTACTO:";
                    workSheet.Cells["B12"].Style.Font.Bold = true;
                    workSheet.Cells["B12"].Style.Font.Size = 11;
                    workSheet.Cells["B12"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                    workSheet.Cells["C12"].Value = cotizacion.Pro_Contacto;
                    workSheet.Cells["C12"].Style.Font.Bold = true;
                    workSheet.Cells["C12"].Style.Font.Size = 11;


                    workSheet.Cells["I12"].Value = "CELULAR:";
                    workSheet.Cells["I12"].Style.Font.Bold = true;
                    workSheet.Cells["I12"].Style.Font.Size = 11;
                    workSheet.Cells["I12"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                    workSheet.Cells["J12"].Value = cotizacion.Prov_Celular;
                    workSheet.Cells["J12"].Style.Font.Bold = true;
                    workSheet.Cells["J12"].Style.Font.Size = 11;

                    workSheet.Cells["B13"].Value = "EMAIL:";
                    workSheet.Cells["B13"].Style.Font.Bold = true;
                    workSheet.Cells["B13"].Style.Font.Size = 11;
                    workSheet.Cells["B13"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                    workSheet.Cells["C13"].Value = cotizacion.Prov_Email;
                    workSheet.Cells["C13"].Style.Font.Bold = true;
                    workSheet.Cells["C13"].Style.Font.Size = 11;

                    ConfigurarTamanioDeCeldas(workSheet);
                    UnirCeldas(workSheet);
                    BordesCeldas(workSheet);
                    PintarCeldas(workSheet);
                    TextoNegrita(workSheet);

                //DETALLE 

                    workSheet.Cells["B14"].Value = "N° ITEM";
                    workSheet.Cells["B14"].Style.Font.Size = 9;
                    workSheet.Cells["B14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["B14"].Style.WrapText = true;
                    workSheet.Cells["B14"].Style.Font.Bold = true;
                    workSheet.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["C14"].Value = "DENOMINACIÓN";
                    workSheet.Cells["C14"].Style.Font.Size = 9;
                    workSheet.Cells["C14"].Style.Font.Bold = true;
                    workSheet.Cells["C14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["C14"].Style.WrapText = true;
                    workSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    workSheet.Cells["D14"].Value = "ESPECIFICACIÓN TECNICA";
                    workSheet.Cells["D14"].Style.Font.Size = 9;
                    workSheet.Cells["D14"].Style.Font.Bold = true;
                    workSheet.Cells["D14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["D14"].Style.WrapText = true;
                    workSheet.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    workSheet.Cells["E14"].Value = "indicaciones y observaciones";
                    workSheet.Cells["E14"].Style.Font.Size = 9;
                    workSheet.Cells["E14"].Style.Font.Bold = true;
                    workSheet.Cells["E14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["E14"].Style.WrapText = true;
                    workSheet.Cells["E14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["F14"].Value = "UM";
                    workSheet.Cells["F14"].Style.Font.Size = 9;
                    workSheet.Cells["F14"].Style.Font.Bold = true;
                    workSheet.Cells["F14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["F14"].Style.WrapText = true;
                    workSheet.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["G14"].Value = "CANT";
                    workSheet.Cells["G14"].Style.Font.Size = 9;
                    workSheet.Cells["G14"].Style.Font.Bold = true;
                    workSheet.Cells["G14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["G14"].Style.WrapText = true;
                    workSheet.Cells["G14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    workSheet.Cells["H14"].Value = "PRECIO UNITARIO S/.(INC. IGV)";
                    workSheet.Cells["H14"].Style.Font.Size = 9;
                    workSheet.Cells["H14"].Style.Font.Bold = true;
                    workSheet.Cells["H14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["H14"].Style.WrapText = true;
                    workSheet.Cells["H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    workSheet.Cells["I14"].Value = "PRECIO TOTAL(S/.) (INC.I.G.V.)";
                    workSheet.Cells["I14"].Style.Font.Size = 9;
                    workSheet.Cells["I14"].Style.Font.Bold = true;
                    workSheet.Cells["I14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["I14"].Style.WrapText = true;
                    workSheet.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    workSheet.Cells["J14"].Value = "PRECIO TOTAL(S/.) (INC.I.G.V.)";
                    workSheet.Cells["J14"].Style.Font.Size = 9;
                    workSheet.Cells["J14"].Style.Font.Bold = true;
                    workSheet.Cells["J14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["J14"].Style.WrapText = true;
                    workSheet.Cells["J14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["K14"].Value = "Nº DE REG. SANITARIO VIGENTE Y LEGIBLE";
                    workSheet.Cells["K14"].Style.Font.Size = 9;
                    workSheet.Cells["K14"].Style.Font.Bold = true;
                    workSheet.Cells["K14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["K14"].Style.WrapText = true;
                    workSheet.Cells["K14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["L14"].Value = "VIGENCIA DEL PRODUCTO COTIZADO";
                    workSheet.Cells["L14"].Style.Font.Size = 9;
                    workSheet.Cells["L14"].Style.Font.Bold = true;
                    workSheet.Cells["L14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["L14"].Style.WrapText = true;
                    workSheet.Cells["L14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["M14"].Value = "VALIDEZ DE OFERTA";
                    workSheet.Cells["M14"].Style.Font.Size = 9;
                    workSheet.Cells["M14"].Style.Font.Bold = true;
                    workSheet.Cells["M14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["M14"].Style.WrapText = true;
                    workSheet.Cells["M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["N14"].Value = "MARCA/PAIS DE PROCEDENCIA";
                    workSheet.Cells["N14"].Style.Font.Size = 9;
                    workSheet.Cells["N14"].Style.Font.Bold = true;
                    workSheet.Cells["N14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["N14"].Style.WrapText = true;
                    workSheet.Cells["N14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["O14"].Value = "PRESENTACION DE PRODUCTO";
                    workSheet.Cells["O14"].Style.Font.Size = 9;
                    workSheet.Cells["O14"].Style.Font.Bold = true;
                    workSheet.Cells["O14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["O14"].Style.WrapText = true;
                    workSheet.Cells["O14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    workSheet.Cells["P14"].Value = "PLAZO DE ENTREGA (Especificar si es Calendarios ó Hábiles)";
                    workSheet.Cells["P14"].Style.Font.Size = 7;
                    workSheet.Cells["P14"].Style.Font.Bold = true;
                    workSheet.Cells["P14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                    workSheet.Cells["P14"].Style.WrapText = true;
                    workSheet.Cells["P14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;



                    int index = 15;

                    foreach (Coti_Formato4_Detalle item in cotizacion.Detalle)
                    {
                        workSheet.Cells["B" + index].Value = index - 14;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["B" + index].Style.Font.Size = 9;
                        workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells["B" + index].Style.Numberformat.Format = "0";

                        workSheet.Cells["C" + index].Value = item.Denominacion;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["C" + index].Style.Font.Size = 9;
                        workSheet.Cells["C" + index].Style.WrapText = true;
                        workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["D" + index].Value = item.EspTecnica;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["D" + index].Style.Font.Size = 9;
                        workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["E" + index].Value = item.IndicadorObs;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["E" + index].Style.Font.Size = 9;
                        workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["F" + index].Value = item.Um;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["F" + index].Style.Font.Size = 9;
                        workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["G" + index].Value = item.Cantidad;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["G" + index].Style.Font.Size = 9;
                        workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells["G" + index].Style.Numberformat.Format = "#,##0";

                        workSheet.Cells["H" + index].Value = item.PreUnitarioIgv;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["H" + index].Style.Font.Size = 9;
                        workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0.00";

                        workSheet.Cells["I" + index].Value = item.PreTotaligv;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["I" + index].Style.Numberformat.Format = "#,##0.00";
                        workSheet.Cells["I" + index].Style.Font.Size = 9;
                        workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["J" + index].Value = item.RnpVigente;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["J" + index].Style.Font.Size = 9;
                        workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["K" + index].Value = item.RegSanitario;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["K" + index].Style.Font.Size = 9;
                        workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                        workSheet.Cells["L" + index].Value = item.VigProducto;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["L" + index].Style.Font.Size = 9;
                        workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["M" + index].Value = item.ValidezOferta;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["M" + index].Style.Font.Size = 9;
                        workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["N" + index].Value = item.MarcapaisProcedencia;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["N" + index].Style.Font.Size = 9;
                        workSheet.Cells["N" + index].Style.WrapText = true;
                        workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        workSheet.Cells["O" + index].Value = item.PresentaProducto;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["O" + index].Style.Font.Size = 9;
                        workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        workSheet.Cells["O" + index].Style.WrapText = true;

                        workSheet.Cells["P" + index].Value = item.PlazaEntrega;
                        workSheet.Row(index).Height = 70.00;
                        workSheet.Cells["P" + index].Style.Font.Size = 9;
                        workSheet.Cells["P" + index].Style.WrapText = true;
                        workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                        index++;
                }
                

                    
                    workSheet.Cells["B" + index + ":J" + index].Merge = true;
                    workSheet.Cells["B" + index].Value = "(*) Los datos solicitados deberan ser llenados en su totalidad (Numero de RUC, numero de Reg. Sanitario, vigencia, marcas, procedencia, plazos de entrega y demas).";
                    workSheet.Cells["B" + index].Style.Font.Size = 11;
                    workSheet.Cells["B" + index].Style.Font.Name = "Batang";


                    index++;
                    workSheet.Cells["B" + index + ":J" + index].Merge = true;
                    workSheet.Cells["B" + index].Value = "(*) la validez de oferta sera congruente respecto al tiempo requerido para procesos de selelcción, según corresponda.";
                    workSheet.Cells["B" + index].Style.Font.Size = 11;
                    workSheet.Cells["B" + index].Style.Font.Name = "Batang";

                    index++;
                    workSheet.Cells["B" + index + ":J" + index].Merge = true;
                    workSheet.Cells["B" + index].Value = "(*) indicar si el producto no es afecto al IGV y arenceles.";
                    workSheet.Cells["B" + index].Style.Font.Size = 11;
                    workSheet.Cells["B" + index].Style.Font.Name = "Batang";

                    index++;
                    workSheet.Cells["B" + index + ":J" + index].Merge = true;
                    workSheet.Cells["B" + index].Value = "Recepcion de cotizaciones hasta 24 horas de remitido el correo de invitacion";
                    workSheet.Cells["B" + index].Style.Font.Size = 11;
                    workSheet.Cells["B" + index].Style.Font.Name = "Batang";


                

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
                   
              
                 workSheet.Column(1).Width = 2.71 + 0.71;
                 workSheet.Column(2).Width = 20.00 + 0.71;
                 workSheet.Column(3).Width = 25.14 + 0.71;
                 workSheet.Column(4).Width = 36.57 + 0.71;
                 workSheet.Column(5).Width = 16.57 + 0.71;
                 workSheet.Column(6).Width = 8.71 + 0.71;
                 workSheet.Column(7).Width = 10.71 + 0.71;
                 workSheet.Column(8).Width = 10.71 + 0.71;
                 workSheet.Column(9).Width = 11.29 + 0.71;
                 workSheet.Column(10).Width = 10.71 + 0.71;
                 workSheet.Column(11).Width = 10.71 + 0.71;
                 workSheet.Column(12).Width = 10.71 + 0.71;
                 workSheet.Column(13).Width = 10.71 + 0.71;
                 workSheet.Column(14).Width = 13.14 + 0.71;
                 workSheet.Column(15).Width = 17.14 + 0.71;
                 workSheet.Column(16).Width = 10.71 + 0.71;



                 workSheet.Row(1).Height = 13.5;
                 workSheet.Row(2).Height = 15.75;
                 workSheet.Row(3).Height = 24.75;
                 workSheet.Row(4).Height = 36.00;
                 workSheet.Row(5).Height = 16.5;
                 workSheet.Row(6).Height = 15.75;
                 workSheet.Row(7).Height = 81.75;
                 workSheet.Row(8).Height = 18;
                 workSheet.Row(9).Height = 18;
                 workSheet.Row(10).Height = 18;
                 workSheet.Row(11).Height = 18;
                 workSheet.Row(12).Height = 18;
                 workSheet.Row(13).Height = 18;
                 workSheet.Row(14).Height = 70.00;
           

                    
        }

            
            private static void UnirCeldas(ExcelWorksheet workSheet)
            {


                 workSheet.Cells["B2:C2"].Merge = true;
                 workSheet.Cells["N2:P3"].Merge = true;
                 workSheet.Cells["N4:P4"].Merge = true;
                 workSheet.Cells["B3:D3"].Merge = true;
                 workSheet.Cells["B4:K4"].Merge = true;
                 workSheet.Cells["B6:C6"].Merge = true;
                 workSheet.Cells["B7:P7"].Merge = true;
                 workSheet.Cells["C8:P8"].Merge = true;
                 workSheet.Cells["C9:P9"].Merge = true;
                 workSheet.Cells["C10:P10"].Merge = true;
                 workSheet.Cells["C11:P11"].Merge = true;
                 workSheet.Cells["C12:H12"].Merge = true;
                 workSheet.Cells["J12:P12"].Merge = true;
                 workSheet.Cells["C13:P13"].Merge = true;

        }

            private static void BordesCeldas(ExcelWorksheet workSheet)
            {
                workSheet.Cells["B2:P7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["N2:P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["N4:P4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                
                workSheet.Cells["C8:P8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C9:P9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C10:P10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["B11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C11:P11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C12:H12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["I12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["J12:P12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["B13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C13:P13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["B14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["D14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["F14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["G14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["H14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["I14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["J14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["K14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["L14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["M14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["N14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["O14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["P14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
            {

                workSheet.Cells["B8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["B9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["B10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["B11"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["B12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["I12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["B13"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                 
                workSheet.Cells["B14:G14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["H14:I14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#008080"));
                workSheet.Cells["J14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00CCFF"));
                workSheet.Cells["K14:P14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#008080"));


        }

            private static void TextoNegrita(ExcelWorksheet workSheet)
            {
                
             }

    }
}
