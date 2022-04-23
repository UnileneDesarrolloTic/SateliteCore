using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato63_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato63_Model cotizacion)
        {
            byte[] file;

            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Breña");

                //ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);


                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                workSheet.Cells["A4"].Value = "DECLARACION JURADA DE COTIZACIÓN DE MATERIAL MEDICO";
                workSheet.Cells["A4"].Style.Font.Size = 15;
                workSheet.Cells["A4"].Style.Font.Name = "Arial";
                workSheet.Cells["A4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A4"].Style.WrapText = true;
                workSheet.Cells["A4"].Style.Font.Bold = true;

                workSheet.Cells["A6"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A6"].Style.Font.Size = 10;
                workSheet.Cells["A6"].Style.Font.Name = "Arial";
                workSheet.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A6"].Style.WrapText = true;
                workSheet.Cells["A6"].Style.Font.Bold = true;


                workSheet.Cells["K6"].Value = "DATOS DEL CONTACTO (REPRESENTANTE DE LA EMPRESA)";
                workSheet.Cells["K6"].Style.Font.Size = 10;
                workSheet.Cells["K6"].Style.Font.Name = "Arial";
                workSheet.Cells["K6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K6"].Style.WrapText = true;
                workSheet.Cells["K6"].Style.Font.Bold = true;


                workSheet.Cells["A7"].Value = "RAZÓN SOCIAL";
                workSheet.Cells["A7"].Style.Font.Size = 10;
                workSheet.Cells["A7"].Style.Font.Name = "Arial";
                workSheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                

                workSheet.Cells["C7"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C7"].Style.Font.Size = 10;
                workSheet.Cells["C7"].Style.Font.Name = "Arial";
                workSheet.Cells["C7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C7"].Style.WrapText = true;

                workSheet.Cells["A8"].Value = "DIRECCIÓN";
                workSheet.Cells["A8"].Style.Font.Size = 10;
                workSheet.Cells["A8"].Style.Font.Name = "Arial";
                workSheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
               

                workSheet.Cells["C8"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["C8"].Style.Font.Size = 10;
                workSheet.Cells["C8"].Style.Font.Name = "Arial";
                workSheet.Cells["C8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C8"].Style.WrapText = true;

                workSheet.Cells["A9"].Value = "TELÉFONO (S)";
                workSheet.Cells["A9"].Style.Font.Size = 10;
                workSheet.Cells["A9"].Style.Font.Name = "Arial";
                workSheet.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                


                workSheet.Cells["C9"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["C9"].Style.Font.Size = 10;
                workSheet.Cells["C9"].Style.Font.Name = "Arial";
                workSheet.Cells["C9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C9"].Style.WrapText = true;


                workSheet.Cells["A10"].Value = "EMAIL";
                workSheet.Cells["A10"].Style.Font.Size = 10;
                workSheet.Cells["A10"].Style.Font.Name = "Arial";
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                

                workSheet.Cells["C10"].Value = cotizacion.Prov_Email;
                workSheet.Cells["C10"].Style.Font.Size = 10;
                workSheet.Cells["C10"].Style.Font.Name = "Arial";
                workSheet.Cells["C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C10"].Style.WrapText = true;


                workSheet.Cells["A11"].Value = "VIGENCIA DE LA OFERTA";
                workSheet.Cells["A11"].Style.Font.Size = 10;
                workSheet.Cells["A11"].Style.Font.Name = "Arial";
                workSheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A11"].Style.WrapText = true;

                workSheet.Cells["C11"].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["C11"].Style.Font.Size = 10;
                workSheet.Cells["C11"].Style.Font.Name = "Arial";
                workSheet.Cells["C11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C11"].Style.WrapText = true;

                workSheet.Cells["F11"].Value = "RUC";
                workSheet.Cells["F11"].Style.Font.Size = 10;
                workSheet.Cells["F11"].Style.Font.Name = "Arial";
                workSheet.Cells["F11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F11"].Style.WrapText = true;

                workSheet.Cells["G11"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["G11"].Style.Font.Size = 10;
                workSheet.Cells["G11"].Style.Font.Name = "Arial";
                workSheet.Cells["G11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G11"].Style.WrapText = true;


                workSheet.Cells["K7"].Value = "NOMBRE";
                workSheet.Cells["K7"].Style.Font.Size = 10;
                workSheet.Cells["K7"].Style.Font.Name = "Arial";   
                workSheet.Cells["K7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["K7"].Style.WrapText = true;

                workSheet.Cells["L7"].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["L7"].Style.Font.Size = 10;
                workSheet.Cells["L7"].Style.Font.Name = "Arial";
                workSheet.Cells["L7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L7"].Style.WrapText = true;

                workSheet.Cells["K8"].Value = "CARGO";
                workSheet.Cells["K8"].Style.Font.Size = 10;
                workSheet.Cells["K8"].Style.Font.Name = "Arial";
                workSheet.Cells["K8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["K8"].Style.WrapText = true;

                workSheet.Cells["L8"].Value = cotizacion.Prov_ContCargo;
                workSheet.Cells["L8"].Style.Font.Size = 10;
                workSheet.Cells["L8"].Style.Font.Name = "Arial";
                workSheet.Cells["L8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L8"].Style.WrapText = true;

                workSheet.Cells["K9"].Value = "CELULAR";
                workSheet.Cells["K9"].Style.Font.Size = 10;
                workSheet.Cells["K9"].Style.Font.Name = "Arial";
                workSheet.Cells["K9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["K9"].Style.WrapText = true;

                workSheet.Cells["L9"].Value = cotizacion.Prov_ContCelular;
                workSheet.Cells["L9"].Style.Font.Size = 10;
                workSheet.Cells["L9"].Style.Font.Name = "Arial";
                workSheet.Cells["L9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L9"].Style.WrapText = true;

                workSheet.Cells["K10"].Value = "RPM / RPC / ENTEL";
                workSheet.Cells["K10"].Style.Font.Size = 10;
                workSheet.Cells["K10"].Style.Font.Name = "Arial";
                workSheet.Cells["K10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["K10"].Style.WrapText = true;

                workSheet.Cells["L10"].Value = cotizacion.Prov_RpmRrc;
                workSheet.Cells["L10"].Style.Font.Size = 10;
                workSheet.Cells["L10"].Style.Font.Name = "Arial";
                workSheet.Cells["L10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L10"].Style.WrapText = true;

                workSheet.Cells["K11"].Value = "EMAIL";
                workSheet.Cells["K11"].Style.Font.Size = 10;
                workSheet.Cells["K11"].Style.Font.Name = "Arial";
                workSheet.Cells["K11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["K11"].Style.WrapText = true;

                workSheet.Cells["L11"].Value = cotizacion.Prov_ContEmail;
                workSheet.Cells["L11"].Style.Font.Size = 10;
                workSheet.Cells["L11"].Style.Font.Name = "Arial";
                workSheet.Cells["L11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L11"].Style.WrapText = true;



                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);

                workSheet.Cells["A14"].Value = "N° ITEM";
                workSheet.Cells["A14"].Style.Font.Size = 9;
                workSheet.Cells["A14"].Style.Font.Name = "Arial";
                workSheet.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A14"].Style.WrapText = true;
                workSheet.Cells["A14"].Style.Font.Bold = true;

                workSheet.Cells["B14"].Value = "DESCRIPCIÓN DEL ÍTEM";
                workSheet.Cells["B14"].Style.Font.Size = 10;
                workSheet.Cells["B14"].Style.Font.Name = "Arial";
                workSheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B14"].Style.WrapText = true;
                workSheet.Cells["B14"].Style.Font.Bold = true;

                workSheet.Cells["F14"].Value = "REQUERIMIENTOS MÍNIMOS";
                workSheet.Cells["F14"].Style.Font.Size = 10;
                workSheet.Cells["F14"].Style.Font.Name = "Arial";
                workSheet.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F14"].Style.WrapText = true;
                workSheet.Cells["F14"].Style.Font.Bold = true;

                workSheet.Cells["H15"].Value = "CALIDAD";
                workSheet.Cells["H15"].Style.Font.Size = 10;
                workSheet.Cells["H15"].Style.Font.Name = "Arial";
                workSheet.Cells["H15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H15"].Style.WrapText = true;
                workSheet.Cells["H15"].Style.Font.Bold = true;



                workSheet.Cells["B16"].Value = "DENOMINACION";
                workSheet.Cells["B16"].Style.Font.Size = 10;
                workSheet.Cells["B16"].Style.Font.Name = "Arial";
                workSheet.Cells["B16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B16"].Style.WrapText = true;
                workSheet.Cells["B16"].Style.Font.Bold = true;

                workSheet.Cells["D16"].Value = "UM";
                workSheet.Cells["D16"].Style.Font.Size = 10;
                workSheet.Cells["D16"].Style.Font.Name = "Arial";
                workSheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D16"].Style.WrapText = true;
                workSheet.Cells["D16"].Style.Font.Bold = true;

                workSheet.Cells["E16"].Value = "REQUERIMIENTO TOTAL";
                workSheet.Cells["E16"].Style.Font.Size = 10;
                workSheet.Cells["E16"].Style.Font.Name = "Arial";
                workSheet.Cells["E16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E16"].Style.WrapText = true;
                workSheet.Cells["E16"].Style.Font.Bold = true;



                workSheet.Cells["F15"].Value = "Pais de Procedencia del Producto";
                workSheet.Cells["F15"].Style.Font.Size = 10;
                workSheet.Cells["F15"].Style.Font.Name = "Arial";
                workSheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F15"].Style.WrapText = true;
                workSheet.Cells["F15"].Style.TextRotation = 90;

                workSheet.Cells["G15"].Value = "Forma de Presentación del Producto";
                workSheet.Cells["G15"].Style.Font.Size = 10;
                workSheet.Cells["G15"].Style.Font.Name = "Arial";
                workSheet.Cells["G15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G15"].Style.WrapText = true;
                workSheet.Cells["G15"].Style.TextRotation = 90;

                workSheet.Cells["H16"].Value = "La vigencia mínima de los materiales  e insumos de laboratorio deberá ser de 18 meses contados a partir de la entrega. (salvo excepción) (Si o No)";
                workSheet.Cells["H16"].Style.Font.Size = 10;
                workSheet.Cells["H16"].Style.Font.Name = "Arial";
                workSheet.Cells["H16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H16"].Style.WrapText = true;
                workSheet.Cells["H16"].Style.TextRotation = 90;

                workSheet.Cells["I16"].Value = "Constancia de Autorización Sanitaria o de estar inscrito como Establecimiento Farmacéutico";
                workSheet.Cells["I16"].Style.Font.Size = 10;
                workSheet.Cells["I16"].Style.Font.Name = "Arial";
                workSheet.Cells["I16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I16"].Style.WrapText = true;
                workSheet.Cells["I16"].Style.TextRotation = 90;

                workSheet.Cells["J16"].Value = "Cumple al 100% con la denominación y descripción del ítem, especificaciones tecnicas y/o terminos de referencia. (Si  ó No)";
                workSheet.Cells["J16"].Style.Font.Size = 10;
                workSheet.Cells["J16"].Style.Font.Name = "Arial";
                workSheet.Cells["J16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J16"].Style.WrapText = true;
                workSheet.Cells["J16"].Style.TextRotation = 90;
                workSheet.Cells["J16"].Style.Font.Bold = true;

                workSheet.Cells["K16"].Value = "Cuenta con Registro Sanitario O Certificado de Registro Sanitario vigente(Si o No)";
                workSheet.Cells["K16"].Style.Font.Size = 10;
                workSheet.Cells["K16"].Style.Font.Name = "Arial";
                workSheet.Cells["K16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K16"].Style.WrapText = true;
                workSheet.Cells["K16"].Style.TextRotation = 90;

                workSheet.Cells["L16"].Value = "Protocolo y/o Certificado de Análisis (En idioma castellano y en original o copia simple).E";
                workSheet.Cells["L16"].Style.Font.Size = 10;
                workSheet.Cells["L16"].Style.Font.Name = "Arial";
                workSheet.Cells["L16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L16"].Style.WrapText = true;
                workSheet.Cells["L16"].Style.TextRotation = 90;

                workSheet.Cells["M16"].Value = "Protocolo y/o Certificado de Análisis (En idioma castellano y en original o copia simple).E";
                workSheet.Cells["M16"].Style.Font.Size = 10;
                workSheet.Cells["M16"].Style.Font.Name = "Arial";
                workSheet.Cells["M16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M16"].Style.WrapText = true;
                workSheet.Cells["M16"].Style.TextRotation = 90;


                workSheet.Cells["N16"].Value = "La oferta incluye ROTULADO: (Si o No)";
                workSheet.Cells["N16"].Style.Font.Size = 10;
                workSheet.Cells["N16"].Style.Font.Name = "Arial";
                workSheet.Cells["N16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N16"].Style.WrapText = true;
                workSheet.Cells["N16"].Style.TextRotation = 90;
                workSheet.Cells["N16"].Style.Font.Bold = true;


                workSheet.Cells["O14"].Value = "Plazo de entrega en dias calendario";
                workSheet.Cells["O14"].Style.Font.Size = 10;
                workSheet.Cells["O14"].Style.Font.Name = "Arial";
                workSheet.Cells["O14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O14"].Style.WrapText = true;
                workSheet.Cells["O14"].Style.TextRotation = 90;

                workSheet.Cells["P14"].Value = "MARCA/ MODELO";
                workSheet.Cells["P14"].Style.Font.Size = 10;
                workSheet.Cells["P14"].Style.Font.Name = "Arial";
                workSheet.Cells["P14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P14"].Style.WrapText = true;
                workSheet.Cells["P14"].Style.TextRotation = 90;


                workSheet.Cells["Q14"].Value = "PROCEDENCIA";
                workSheet.Cells["Q14"].Style.Font.Size = 10;
                workSheet.Cells["Q14"].Style.Font.Name = "Arial";
                workSheet.Cells["Q14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q14"].Style.WrapText = true;
                workSheet.Cells["Q14"].Style.TextRotation = 90;

                workSheet.Cells["R14"].Value = "FORMA DE PAGO";
                workSheet.Cells["R14"].Style.Font.Size = 10;
                workSheet.Cells["R14"].Style.Font.Name = "Arial";
                workSheet.Cells["R14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R14"].Style.WrapText = true;
                workSheet.Cells["R14"].Style.TextRotation = 90;

                workSheet.Cells["S14"].Value = "PRECIO UNITARIO A OFERTAR S /.";
                workSheet.Cells["S14"].Style.Font.Size = 10;
                workSheet.Cells["S14"].Style.Font.Name = "Arial";
                workSheet.Cells["S14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S14"].Style.WrapText = true;
                workSheet.Cells["S14"].Style.TextRotation = 90;
                workSheet.Cells["S14"].Style.Font.Bold = true;

                workSheet.Cells["T14"].Value = "VALOR TOTAL A OFERTAR S /.";
                workSheet.Cells["T14"].Style.Font.Size = 10;
                workSheet.Cells["T14"].Style.Font.Name = "Arial";
                workSheet.Cells["T14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T14"].Style.WrapText = true;
                workSheet.Cells["T14"].Style.TextRotation = 90;
                workSheet.Cells["T14"].Style.Font.Bold = true;

                int index = 19;

                foreach (Coti_Formato63_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Row(index).Height = 30.75;

                    workSheet.Cells["A" + index].Value = index - 18;
                    workSheet.Cells["A" + index].Style.Font.Size = 8;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["A" + index].Style.WrapText = true;

                    workSheet.Cells["B" + index + ":C" + index].Merge = true;
                    workSheet.Cells["B" + index + ":C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index + ":C" + index].Value = item.Denominacion;
                    workSheet.Cells["B" + index + ":C" + index].Style.Font.Size = 8;
                    workSheet.Cells["B" + index + ":C" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index + ":C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index + ":C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index + ":C" + index].Style.WrapText = true;

                    workSheet.Cells["D" + index].Value = item.Um;
                    workSheet.Cells["D" + index].Style.Font.Size = 8;
                    workSheet.Cells["D" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;

                    workSheet.Cells["E" + index].Value = item.ReqTotal;
                    workSheet.Cells["E" + index].Style.Font.Size = 8;
                    workSheet.Cells["E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;


                    workSheet.Cells["F" + index].Value = item.PaisProcedencia;
                    workSheet.Cells["F" + index].Style.Font.Size = 8;
                    workSheet.Cells["F" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;


                    workSheet.Cells["G" + index].Value = item.Presentacion;
                    workSheet.Cells["G" + index].Style.Font.Size = 7;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;


                    workSheet.Cells["H" + index].Value = item.VigMaterial;
                    workSheet.Cells["H" + index].Style.Font.Size = 8;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;

                    workSheet.Cells["I" + index].Value = item.AutoSanitaria;
                    workSheet.Cells["I" + index].Style.Font.Size = 8;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Cells["J" + index].Value = item.CumpTerminos;
                    workSheet.Cells["J" + index].Style.Font.Size = 8;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;

                    workSheet.Cells["K" + index].Value = item.RegSanitario;
                    workSheet.Cells["K" + index].Style.Font.Size = 8;
                    workSheet.Cells["K" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;

                    workSheet.Cells["L" + index].Value = item.CertAnalisis;
                    workSheet.Cells["L" + index].Style.Font.Size = 8;
                    workSheet.Cells["L" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;



                    workSheet.Cells["M" + index].Value = item.BuenasPracticas;
                    workSheet.Cells["M" + index].Style.Font.Size = 8;
                    workSheet.Cells["M" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    workSheet.Cells["N" + index].Value = item.Rotulado;
                    workSheet.Cells["N" + index].Style.Font.Size = 8;
                    workSheet.Cells["N" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true; 


                    workSheet.Cells["O" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["O" + index].Style.Font.Size = 8;
                    workSheet.Cells["O" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;

                    workSheet.Cells["P" + index].Value = item.MarcaModelo;
                    workSheet.Cells["P" + index].Style.Font.Size = 8;
                    workSheet.Cells["P" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["P" + index].Style.WrapText = true;

                    workSheet.Cells["Q" + index].Value = item.Procedencia;
                    workSheet.Cells["Q" + index].Style.Font.Size = 8;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;

                    workSheet.Cells["R" + index].Value = item.FormaPago;
                    workSheet.Cells["R" + index].Style.Font.Size = 8;
                    workSheet.Cells["R" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["R" + index].Style.WrapText = true;

                    workSheet.Cells["S" + index].Value = item.PreUnitario;
                    workSheet.Cells["S" + index].Style.Font.Size = 8;
                    workSheet.Cells["S" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["S" + index].Style.WrapText = true;
                    workSheet.Cells["S" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["T" + index].Value = item.ValorTotal;
                    workSheet.Cells["T" + index].Style.Font.Size = 8;
                    workSheet.Cells["T" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["T" + index].Style.WrapText = true;
                    workSheet.Cells["T" + index].Style.Numberformat.Format = "#,##0.00";


                    index++;


                }

                index++;

                
                workSheet.Row(index).Height = 15.75;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "CONSIDERACIONES GENERALES: ";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";

                index++;

                workSheet.Row(index).Height = 14.25;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "1.- El precio de Mercado será a todo costo, es decir, deberá incluir todos los tributos (incluido el I.G.V.), seguros, transportes, inspecciones";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Row(index).Height = 14.25;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "2.-De ser el caso; indicar si su Precio de Mercado está exento del I.G.V. y aranceles.";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Row(index).Height = 34.5;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "3.- Excepcionalmente, para los productos que por sus propiedades biológicas, físicas y químicas no puedan cumplir con la vigencia mínima establecida, deben INDICAR en el respectivo recuadro, la palabra 'EXCEPCION'," +
                    " de ser el caso vigencias menores, siempre que éstas no sean inferiores al 80%  del tiempo de vida útil especificado para el producto y declarado por el fabricante.";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;

                workSheet.Row(index).Height = 29.25;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "4.- Cabe precisar que frente a un reclamo, queja u observación de los materiales e insumos de laboratotio, queda en potestad del INSN-BREÑA el sometimiento al control de calidad respectivo, " +
                    "cuyo costo será asumido enteramente por el proveedor, en caso el resultado sea no conforme. ";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 14.25;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Deberán considerar el rotulado de acuerdo a los TDR";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 14.25;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Lugar de Entrega:";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 19.5;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Av. Brasil 600 - Breña - Lima (Almacén del INSN - BREÑA)";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 14.25;
                workSheet.Cells["A" + index + ":T" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "NOTA: SU COTIZACIÓN DEBERÁ DE CONSIGNAR ADEMÁS: LOGOTIPO (si fuera el caso), y ser firmado por el representante legal de la empresa.";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";

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
            workSheet.Column(1).Width = 4.29 + 0.71;
            workSheet.Column(2).Width = 10.43 + 0.71;
            workSheet.Column(3).Width = 33.43 + 0.71;
            workSheet.Column(4).Width = 5.86 + 0.71;
            workSheet.Column(5).Width = 12.29 + 0.71;
            workSheet.Column(6).Width = 6.14 + 0.71;
            workSheet.Column(7).Width = 4.71 + 0.71;
            workSheet.Column(8).Width = 12.57 + 0.71;
            workSheet.Column(9).Width = 10.71 + 0.71;
            workSheet.Column(10).Width = 13.14 + 0.71;
            workSheet.Column(11).Width = 11.43 + 0.71;
            workSheet.Column(12).Width = 9.00 + 0.71;
            workSheet.Column(13).Width = 8.86 + 0.71;
            workSheet.Column(14).Width = 8.86 + 0.71;
            workSheet.Column(15).Width = 6.86 + 0.71;
            workSheet.Column(16).Width = 6.86 + 0.71;
            workSheet.Column(17).Width = 6.86 + 0.71;
            workSheet.Column(18).Width = 6.86 + 0.71;
            workSheet.Column(19).Width = 11.57 + 0.71;
            workSheet.Column(20).Width = 13.00 + 0.71;
            workSheet.Column(21).Width = 0.42 + 0.71;
            


            workSheet.Row(1).Height = 12.75;
            workSheet.Row(2).Height = 12.75;
            workSheet.Row(3).Height = 12.75;
            workSheet.Row(4).Height = 19.50;
            workSheet.Row(5).Height = 13.50;
            workSheet.Row(6).Height = 24.00;
            workSheet.Row(7).Height = 26.25;
            workSheet.Row(8).Height = 20.00;
            workSheet.Row(9).Height = 25.50;
            workSheet.Row(10).Height = 31.5;
            workSheet.Row(11).Height = 31.50;
            workSheet.Row(12).Height = 12.75;
            workSheet.Row(13).Height = 5.25;

            workSheet.Row(14).Height = 18.00;
            workSheet.Row(15).Height = 18.00;
            workSheet.Row(16).Height = 60.00;
            workSheet.Row(17).Height = 60.00;
            workSheet.Row(18).Height = 84.75;
            
        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A4:T4"].Merge = true;
            workSheet.Cells["A6:I6"].Merge = true;
            workSheet.Cells["K6:U6"].Merge = true;
            workSheet.Cells["A7:B7"].Merge = true;
            workSheet.Cells["C7:I7"].Merge = true;
            workSheet.Cells["A8:B8"].Merge = true;
            workSheet.Cells["C8:I8"].Merge = true;
            workSheet.Cells["A9:B9"].Merge = true;
            workSheet.Cells["C9:I9"].Merge = true;
            workSheet.Cells["A10:B10"].Merge = true;
            workSheet.Cells["C10:I10"].Merge = true;
            workSheet.Cells["A11:B11"].Merge = true;
            workSheet.Cells["C11:E11"].Merge = true;
            workSheet.Cells["G11:I11"].Merge = true;
            workSheet.Cells["L7:U7"].Merge = true;
            workSheet.Cells["L8:U8"].Merge = true;
            workSheet.Cells["L9:U9"].Merge = true;
            workSheet.Cells["L10:U10"].Merge = true;
            workSheet.Cells["L11:U11"].Merge = true;

            //DETALLE
            workSheet.Cells["A14:A18"].Merge = true;
            workSheet.Cells["B14:E15"].Merge = true;

            workSheet.Cells["B16:C18"].Merge = true;
            workSheet.Cells["D16:D18"].Merge = true;
            workSheet.Cells["E16:E18"].Merge = true;


            workSheet.Cells["F14:N14"].Merge = true;
            workSheet.Cells["H15:N15"].Merge = true;


            workSheet.Cells["F15:F18"].Merge = true;
            workSheet.Cells["G15:G18"].Merge = true;

            workSheet.Cells["H16:H18"].Merge = true;
            workSheet.Cells["I16:I18"].Merge = true;
            workSheet.Cells["J16:J18"].Merge = true;
            workSheet.Cells["K16:K18"].Merge = true;
            workSheet.Cells["L16:L18"].Merge = true;
            workSheet.Cells["M16:M18"].Merge = true;
            workSheet.Cells["N16:N18"].Merge = true;
            workSheet.Cells["O14:O18"].Merge = true;
            workSheet.Cells["P14:P18"].Merge = true;
            workSheet.Cells["Q14:Q18"].Merge = true;
            workSheet.Cells["R14:R18"].Merge = true;
            workSheet.Cells["S14:S18"].Merge = true;
            workSheet.Cells["T14:T18"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A6:I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K6:U6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A7:B7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C7:I7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A8:B8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C8:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A9:B9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C9:I9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A10:B10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C10:I10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A11:B11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C11:E11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G11:I11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L7:U7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L8:U8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L9:U9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L10:U10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L11:U11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["A14:A18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B14:E15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["B16:C18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D16:D18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E16:E18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F14:N14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H15:N15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F15:F18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["G15:G18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H16:H18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I16:I18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J16:J18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K16:K18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L16:L18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["M16:M18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N16:N18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O14:O18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P14:P18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q14:Q18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R14:R18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S14:S18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["T14:T18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


        }
    }
}
