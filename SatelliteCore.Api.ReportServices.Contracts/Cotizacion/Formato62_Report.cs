using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato62_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato62_Model cotizacion)
        {
            byte[] file;

            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Almenara");

                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 3, 9);
                imagenUnilene.SetSize(197, 99);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Name = "Calibri";

                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                workSheet.Cells["A2"].Value = "UNIDAD DE PROGRAMACIÓN";
                workSheet.Cells["A2"].Style.Font.Size = 12;
                workSheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A2"].Style.Font.Bold = true;


                workSheet.Cells["A3"].Value = "OFICINA DE ABASTECIMIENTO Y CONTROL PATRIMONIAL";
                workSheet.Cells["A3"].Style.Font.Size = 12;
                workSheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A3"].Style.Font.Bold = true;

                workSheet.Cells["A4"].Value = "RED ASISTENCIAL REBAGLIATI";
                workSheet.Cells["A4"].Style.Font.Size = 12;
                workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A4"].Style.Font.Bold = true;

                workSheet.Cells["D2"].Value = "SOLICITUD  DE  COTIZACIÓN  DE  BIENES";
                workSheet.Cells["D2"].Style.Font.Size = 24;
                workSheet.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D2"].Style.Font.Bold = true;
                workSheet.Cells["D2"].Style.WrapText = true;


                workSheet.Cells["A6"].Value = "Señores \n" +
                    "ESSALUD R \n"+
                    "De mi consideración: \n" +
                    "En respuesta a la solicitud de cotización sobre la 'ADQUISICIÓN................................................................................................................................... HNERM',  y después de haber analizadolas espcificaciones técnicas del mencionado requerimiento, las mismas que acepto en todos los extremos, indico que cumplo con TODOS los requerimientos solicitados. \n" +
                    "Asimismo, declaro que las características técnicas y/o servicios cotizados por mi representada se ajustan a lo requerido por su Entidad. En tal sentido, indico que el costo total por lo requerido es la que detallo a continuación:";
                workSheet.Cells["A6"].Style.Font.Size = 12;
                workSheet.Cells["A6"].Style.Font.Name = "Arial";
                workSheet.Cells["A6"].Style.Font.Bold = true;
                workSheet.Cells["A6"].Style.WrapText = true;
                workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                workSheet.Cells["A7"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A7"].Style.Font.Size = 14;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A7"].Style.Font.Bold = true;
                workSheet.Cells["A7"].Style.WrapText = true;


                workSheet.Cells["A8"].Value = "RAZON SOCIAL";
                workSheet.Cells["A8"].Style.Font.Size = 11;
                workSheet.Cells["A8"].Style.Font.Bold = true;
                workSheet.Cells["A8"].Style.WrapText = true;

                workSheet.Cells["C8"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C8"].Style.Font.Size = 11;
                workSheet.Cells["C8"].Style.Font.Bold = true;
                workSheet.Cells["C8"].Style.WrapText = true;

                workSheet.Cells["J8"].Value = "TELEFONO(S)";
                workSheet.Cells["J8"].Style.Font.Size = 11;
                workSheet.Cells["J8"].Style.Font.Bold = true;
                workSheet.Cells["J8"].Style.WrapText = true;

                workSheet.Cells["K8"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["K8"].Style.Font.Size = 11;
                workSheet.Cells["K8"].Style.Font.Bold = true;
                workSheet.Cells["K8"].Style.WrapText = true;


                workSheet.Cells["M8"].Value = "CONTACTO";
                workSheet.Cells["M8"].Style.Font.Size = 11;
                workSheet.Cells["M8"].Style.Font.Bold = true;
                workSheet.Cells["M8"].Style.WrapText = true;


                workSheet.Cells["O8"].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["O8"].Style.Font.Size = 11;
                workSheet.Cells["O8"].Style.Font.Bold = true;
                workSheet.Cells["O8"].Style.WrapText = true;


                workSheet.Cells["A9"].Value = "RUC";
                workSheet.Cells["A9"].Style.Font.Size = 11; 
                workSheet.Cells["A9"].Style.Font.Bold = true;
                workSheet.Cells["A9"].Style.WrapText = true;

                workSheet.Cells["C9"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["C9"].Style.Font.Size = 11;
                workSheet.Cells["C9"].Style.Font.Bold = true;
                workSheet.Cells["C9"].Style.WrapText = true;

                workSheet.Cells["J9"].Value = "FAX";
                workSheet.Cells["J9"].Style.Font.Size = 11;               
                workSheet.Cells["J9"].Style.Font.Bold = true;
                workSheet.Cells["J9"].Style.WrapText = true;

                workSheet.Cells["K9"].Value = cotizacion.Prov_Fax;
                workSheet.Cells["K9"].Style.Font.Size = 11;
                workSheet.Cells["K9"].Style.Font.Bold = true;
                workSheet.Cells["K9"].Style.WrapText = true;


                workSheet.Cells["M9"].Value = "CARGO";
                workSheet.Cells["M9"].Style.Font.Size = 11;
                workSheet.Cells["M9"].Style.Font.Bold = true;
                workSheet.Cells["M9"].Style.WrapText = true;

                workSheet.Cells["O9"].Value = cotizacion.Prov_Cargo;
                workSheet.Cells["O9"].Style.Font.Size = 11;
                workSheet.Cells["O9"].Style.Font.Bold = true;
                workSheet.Cells["O9"].Style.WrapText = true;



                workSheet.Cells["A10"].Value = "DIRECCION";
                workSheet.Cells["A10"].Style.Font.Size = 11;
                workSheet.Cells["A10"].Style.Font.Bold = true;
                workSheet.Cells["A10"].Style.WrapText = true;

                workSheet.Cells["C10"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["C10"].Style.Font.Size = 11;
                workSheet.Cells["C10"].Style.Font.Bold = true;
                workSheet.Cells["C10"].Style.WrapText = true;


                workSheet.Cells["J10"].Value = "VIGENCIA DE OFERTA";
                workSheet.Cells["J10"].Style.Font.Size = 11;
                workSheet.Cells["J10"].Style.Font.Bold = true;
                workSheet.Cells["J10"].Style.WrapText = true;

                workSheet.Cells["K10"].Value = cotizacion.Vig_Oferta;
                workSheet.Cells["K10"].Style.Font.Size = 11;
                workSheet.Cells["K10"].Style.Font.Bold = true;
                workSheet.Cells["K10"].Style.WrapText = true;


                workSheet.Cells["M10"].Value = "TELEFONO";
                workSheet.Cells["M10"].Style.Font.Size = 11;
                workSheet.Cells["M10"].Style.Font.Bold = true;
                workSheet.Cells["M10"].Style.WrapText = true;

                workSheet.Cells["O10"].Value = cotizacion.Prov_Telefono2;
                workSheet.Cells["O10"].Style.Font.Size = 11;
                workSheet.Cells["O10"].Style.Font.Bold = true;
                workSheet.Cells["O10"].Style.WrapText = true;

                workSheet.Cells["A11"].Value = "E-MAIL";
                workSheet.Cells["A11"].Style.Font.Size = 11;
                workSheet.Cells["A11"].Style.Font.Bold = true;
                workSheet.Cells["A11"].Style.WrapText = true;

                workSheet.Cells["C11"].Value = cotizacion.Prov_Email;
                workSheet.Cells["C11"].Style.Font.Size = 11;
                workSheet.Cells["C11"].Style.Font.Bold = true;
                workSheet.Cells["C11"].Style.WrapText = true;

                workSheet.Cells["J11"].Value = "FECHA DE COTIZACIÓN";
                workSheet.Cells["J11"].Style.Font.Size = 11;
                workSheet.Cells["J11"].Style.Font.Bold = true;
                workSheet.Cells["J11"].Style.WrapText = true;

                workSheet.Cells["K11"].Value = cotizacion.Fecha_Cotizacion;
                workSheet.Cells["K11"].Style.Font.Size = 11;
                workSheet.Cells["K11"].Style.Font.Bold = true;
                workSheet.Cells["K11"].Style.WrapText = true;
                workSheet.Cells["K11"].Style.Numberformat.Format = "dd/MM/yyyy";


                workSheet.Cells["M11"].Value = "CELULAR";
                workSheet.Cells["M11"].Style.Font.Size = 11;
                workSheet.Cells["M11"].Style.Font.Bold = true;
                workSheet.Cells["M11"].Style.WrapText = true;

                workSheet.Cells["O11"].Value = cotizacion.Prov_Celular;
                workSheet.Cells["O11"].Style.Font.Size = 11;
                workSheet.Cells["O11"].Style.Font.Bold = true;
                workSheet.Cells["O11"].Style.WrapText = true;


                workSheet.Cells["A12"].Value = "N° COTIZACIÓN";
                workSheet.Cells["A12"].Style.Font.Size = 11;
                workSheet.Cells["A12"].Style.Font.Bold = true;
                workSheet.Cells["A12"].Style.WrapText = true;

                workSheet.Cells["C12"].Value = cotizacion.Prov_NroCotizacion;
                workSheet.Cells["C12"].Style.Font.Size = 11;
                workSheet.Cells["C12"].Style.Font.Bold = true;
                workSheet.Cells["C12"].Style.WrapText = true;

                workSheet.Cells["D12"].Value = "DATOS ADICIONALES";
                workSheet.Cells["D12"].Style.Font.Size = 11;
                workSheet.Cells["D12"].Style.Font.Bold = true;
                workSheet.Cells["D12"].Style.WrapText = true;

                workSheet.Cells["G12"].Value = cotizacion.Prov_DatosAdicionales;
                workSheet.Cells["G12"].Style.Font.Size = 11;
                workSheet.Cells["G12"].Style.Font.Bold = true;
                workSheet.Cells["G12"].Style.WrapText = true;


                workSheet.Cells["A14"].Value = "DESCRIPCIÓN DEL ÍTEM";
                workSheet.Cells["A14"].Style.Font.Size = 11;
                workSheet.Cells["A14"].Style.WrapText = true;
                workSheet.Cells["A14"].Style.Font.Bold = true;
                workSheet.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["F14"].Value = "DESCRIPCIÓN DEL ÍTEM";
                workSheet.Cells["F14"].Style.Font.Size = 11;
                workSheet.Cells["F14"].Style.WrapText = true;
                workSheet.Cells["F14"].Style.Font.Bold = true;
                workSheet.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;




                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);
                TextoNegrita(workSheet);
                PintarCeldas(workSheet);

                //DETALLE
                workSheet.Cells["A15"].Value = "N° ITEM";
                workSheet.Cells["A15"].Style.Font.Size = 11;
                workSheet.Cells["A15"].Style.WrapText = true;
                workSheet.Cells["A15"].Style.Font.Bold = true;
                workSheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["F15"].Value = "CALIDAD";
                workSheet.Cells["F15"].Style.Font.Size = 12;
                workSheet.Cells["F15"].Style.WrapText = true;
                workSheet.Cells["F15"].Style.Font.Bold = true;
                workSheet.Cells["F15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                workSheet.Cells["B16"].Value = "CODIGO SAP";
                workSheet.Cells["B16"].Style.Font.Size = 12;
                workSheet.Cells["B16"].Style.WrapText = true;
                workSheet.Cells["B16"].Style.Font.Bold = true;
                workSheet.Cells["B16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                workSheet.Cells["C16"].Value = "DENOMINACION O DESCRIPCIÓN";
                workSheet.Cells["C16"].Style.Font.Size = 12;
                workSheet.Cells["C16"].Style.WrapText = true;
                workSheet.Cells["C16"].Style.Font.Bold = true;
                workSheet.Cells["C16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["G16"].Value = "CALIDAD";
                workSheet.Cells["G16"].Style.Font.Size = 12;
                workSheet.Cells["G16"].Style.WrapText = true;
                workSheet.Cells["G16"].Style.Font.Bold = true;
                workSheet.Cells["G16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                workSheet.Cells["I16"].Value = "UNIDAD DE MEDIDA";
                workSheet.Cells["I16"].Style.Font.Size = 12;
                workSheet.Cells["I16"].Style.WrapText = true;
                workSheet.Cells["I16"].Style.Font.Bold = true;
                workSheet.Cells["I16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["J16"].Value = "MODELO";
                workSheet.Cells["J16"].Style.Font.Size = 12;
                workSheet.Cells["J16"].Style.WrapText = true;
                workSheet.Cells["J16"].Style.Font.Bold = true;
                workSheet.Cells["J16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["K16"].Value = "MARCA";
                workSheet.Cells["K16"].Style.Font.Size = 12;
                workSheet.Cells["K16"].Style.WrapText = true;
                workSheet.Cells["K16"].Style.Font.Bold = true;
                workSheet.Cells["K16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["L16"].Value = "PROCEDENCIA";
                workSheet.Cells["L16"].Style.Font.Size = 12;
                workSheet.Cells["L16"].Style.WrapText = true;
                workSheet.Cells["L16"].Style.Font.Bold = true;
                workSheet.Cells["L16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["N16"].Value = "PRESENTACION";
                workSheet.Cells["N16"].Style.Font.Size = 12;
                workSheet.Cells["N16"].Style.WrapText = true;
                workSheet.Cells["N16"].Style.Font.Bold = true;
                workSheet.Cells["N16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["O16"].Value = "VALIDEZ DE OFERTA (Días calendario)";
                workSheet.Cells["O16"].Style.Font.Size = 12;
                workSheet.Cells["O16"].Style.WrapText = true;
                workSheet.Cells["O16"].Style.Font.Bold = true;
                workSheet.Cells["O16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["P16"].Value = "Cumple al 100% con la denominación y descripción del ítem (Si ó No) ";
                workSheet.Cells["P16"].Style.Font.Size = 12;
                workSheet.Cells["P16"].Style.WrapText = true;
                workSheet.Cells["P16"].Style.Font.Bold = true;
                workSheet.Cells["P16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["Q16"].Value = "La vigencia del producto deberá ser mayor a 12 meses";
                workSheet.Cells["Q16"].Style.Font.Size = 12;
                workSheet.Cells["Q16"].Style.WrapText = true;
                workSheet.Cells["Q16"].Style.Font.Bold = true;
                workSheet.Cells["Q16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                workSheet.Cells["R16"].Value = "Capacidad de Atención al 100% de lo solicitado (Si o No)";
                workSheet.Cells["R16"].Style.Font.Size = 12;
                workSheet.Cells["R16"].Style.WrapText = true;
                workSheet.Cells["R16"].Style.Font.Bold = true;
                workSheet.Cells["R16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                workSheet.Cells["S16"].Value = "Plazo de Entrega en Días Calendarios";
                workSheet.Cells["S16"].Style.Font.Size = 12;
                workSheet.Cells["S16"].Style.WrapText = true;
                workSheet.Cells["S16"].Style.Font.Bold = true;
                workSheet.Cells["S16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["T14"].Value = "PRECIO UNITARIO A OFERTAR /.(INCL.IGV)";
                workSheet.Cells["T14"].Style.Font.Size = 12;
                workSheet.Cells["T14"].Style.WrapText = true;
                workSheet.Cells["T14"].Style.Font.Bold = true;
                workSheet.Cells["T14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["V14"].Value = "PRECIO UNITARIO A OFERTAR S/.";
                workSheet.Cells["V14"].Style.Font.Size = 12;
                workSheet.Cells["V14"].Style.WrapText = true;
                workSheet.Cells["V14"].Style.Font.Bold = true;
                workSheet.Cells["V14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int index = 19;

                foreach (Coti_Formato62_Detalle item in cotizacion.Detalle)
                {

                    workSheet.Cells["A" + index].Value = index - 18;
                    workSheet.Cells["A" + index].Style.Font.Size = 14;
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                  
                    workSheet.Cells["B" + index].Value = item.CodigoSap;
                    workSheet.Cells["B" + index].Style.Font.Size = 14;
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index + ":F" + index].Merge = true;
                    workSheet.Cells["C" + index].Value = item.Denominacion;
                    workSheet.Cells["C" + index].Style.Font.Size = 14;
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index + ":H" + index].Merge = true;
                    workSheet.Cells["G" + index].Value = item.Cantidad;
                    workSheet.Cells["G" + index].Style.Font.Size = 14;
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index + ":H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["I" + index].Value = item.UndMedida;
                    workSheet.Cells["I" + index].Style.Font.Size = 14;
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    workSheet.Cells["J" + index].Value = item.Modelo;
                    workSheet.Cells["J" + index].Style.Font.Size = 14;
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = item.Marca;
                    workSheet.Cells["K" + index].Style.Font.Size = 14;
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + index + ":M" + index].Merge = true;
                    workSheet.Cells["L" + index].Value = item.Procedencia;
                    workSheet.Cells["L" + index].Style.Font.Size = 14;
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index + ":M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    
                    workSheet.Cells["N" + index].Value = item.Presentacion;
                    workSheet.Cells["N" + index].Style.Font.Size = 14;
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    workSheet.Cells["O" + index].Value = item.Voferta;
                    workSheet.Cells["O" + index].Style.Font.Size = 14;
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + index].Value = item.CumpDemoninacion;
                    workSheet.Cells["P" + index].Style.Font.Size = 14;
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["Q" + index].Value = item.VigProducto;
                    workSheet.Cells["Q" + index].Style.Font.Size = 14;
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["R" + index].Value = item.CapAtencion;
                    workSheet.Cells["R" + index].Style.Font.Size = 14;
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["S" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["S" + index].Style.Font.Size = 14;
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["T" + index + ":U" + index].Merge = true;
                    workSheet.Cells["T" + index].Value = item.PreUnitario;
                    workSheet.Cells["T" + index].Style.Font.Size = 14;
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["T" + index + ":U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["T" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["V" + index].Value = item.ValorTotal;
                    workSheet.Cells["V" + index].Style.Font.Size = 14;
                    workSheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["V" + index].Style.Numberformat.Format = "#,##0.00";

                    index++;
                }

                index++;
                workSheet.Cells["G" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["G" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["G" + index + ":I" + index].Merge = true;
                workSheet.Cells["G"+ index ].Value = "Marcar con 'X'";
                workSheet.Cells["G" + index].Style.Font.Size = 12;
                workSheet.Cells["G" + index].Style.WrapText = true;
                workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G" + index + ":I" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
                workSheet.Cells["G" + index + ":I" + index].Style.Font.Bold = true;


                //CONSIDERACIONES GENERALES
                workSheet.Cells["K" + index + ":V" + (index+8)].Merge = true;
                workSheet.Cells["K" + index + ":V" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["K" + index + ":V" + (index + 8)].Style.Font.Bold = true;
                workSheet.Cells["K" + index].Value = "CONSIDERACIONES GENERALES \n" +
                "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye todos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso, los costos laborales conforme la legislación vigente, así como culaquier u otro concepto que pueda tener incidencia sobre el costo del bien a contratar, excepto la de aquellos proveedores que gocen de alguna exoneración legal, no incluirán en el precio en el precio de su oferta los tributos respectivos. \n" +
                "Asimismo, declaro bajo juramento que, mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, conforme lo establece el artículo 11 del TUO de la Ley de Contrataciones del Estado, aprobado por D.S. Nº 082-2019-EF \n" +
                "IMPORTANTE \n" +
                "1. Todos los ítems deberán de remitirse con sus respectivas especificaciones técnicas del producto, dentro de la cotización del proveedor o por separado. \n" +
                "2. Se solicita adjuntar constancia RNP. \n" +
                "3. Deberán considerar el rotulado de 'EsSalud prohibida su ventA'. \n"+
                "REMITIR LA PRESENTE EN HOJA MEMBRETADA Y FIRMADA";
                workSheet.Cells["K" + index].Style.Font.Size = 12;
                workSheet.Cells["K" + index].Style.WrapText = true;
                workSheet.Cells["K" + index].Style.Font.Name = "Arial";
                workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


              
        



                index++;
                workSheet.Cells["G" + index].Value = "SI";
                workSheet.Cells["G" + index].Style.Font.Size = 12;
                workSheet.Cells["G" + index].Style.WrapText = true;
                workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H" + index].Value = "NO";
                workSheet.Cells["H" + index].Style.Font.Size = 12;
                workSheet.Cells["H" + index].Style.WrapText = true;
                workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                workSheet.Cells["I" + index].Value = "OBSERVACIÓN(*)";
                workSheet.Cells["I" + index].Style.Font.Size = 12;
                workSheet.Cells["I" + index].Style.WrapText = true;
                workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Cells["A" + index].Value = "DECLARO EL CUMPLIMIENTO DE LAS ESPECIFICACIONES TECNICAS";
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":F" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":F" + index].Merge = true;
                workSheet.Cells["A" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G" + index].Style.Font.Bold = true;
                workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["H" + index].Style.Font.Bold = true;
                workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["I" + index].Style.Font.Bold = true;
                workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                index++;
                workSheet.Cells["A" + index].Value = "CUENTO CON REGISTRO NACIONAL DE PROVEEDORES (RNP)";
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":F" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":F" + index].Merge = true;
                workSheet.Cells["A" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G" + index].Style.Font.Bold = true;
                workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["H" + index].Style.Font.Bold = true;
                workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["I" + index].Style.Font.Bold = true;
                workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                index++;
                workSheet.Cells["A" + index].Value = "GARANTÍA MINIMA (indicar meses o años) ";
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":F" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":F" + index].Merge = true;
                workSheet.Cells["A" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G" + index].Value = cotizacion.Prov_GarantiaMinima;
                workSheet.Cells["G" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["G" + index + ":I" + index].Merge = true;
                workSheet.Cells["G" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;
                workSheet.Cells["A" + index].Value = "FORMA DE PAGO";
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":F" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":F" + index].Merge = true;
                workSheet.Cells["A" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G" + index].Value = cotizacion.Prov_FormaPago;
                workSheet.Cells["G" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["G" + index + ":I" + index].Merge = true;
                workSheet.Cells["G" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Cells["A" + index].Value = "VIGENCIA DE LA COTIZACIÓN";
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":F" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":F" + index].Merge = true;
                workSheet.Cells["A" + index + ":F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                workSheet.Cells["G" + index].Value = cotizacion.Prov_Vigcotizacion;
                workSheet.Cells["G" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["G" + index + ":I" + index].Merge = true;
                workSheet.Cells["G" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;
                workSheet.Row(index).Height =73.50;
                workSheet.Cells["A" + index].Value = "MEJORAS A OFRECER:";
                workSheet.Cells["A" + index].Style.Font.Size = 12;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":B" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
                workSheet.Cells["A" + index + ":B" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.WrapText = true;


                workSheet.Cells["C" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":I" + index].Merge = true;
                workSheet.Cells["C" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Row(index).Height = 66.75;
                workSheet.Cells["A" + index].Value = "(*)OBSERVACIONES";
                workSheet.Cells["A" + index].Style.Font.Size = 9;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":B" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":B" + index].Merge = true;
                workSheet.Cells["A" + index + ":B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.WrapText = true;


                workSheet.Cells["C" + index + ":I" + index].Style.Font.Bold = true;
                workSheet.Cells["C" + index + ":I" + index].Merge = true;
                workSheet.Cells["C" + index + ":I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                index++;
                workSheet.Cells["C" + index + ":F" + (index+5)].Merge = true;
                workSheet.Cells["C" + index + ":F" + (index + 5)].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                workSheet.Cells["C" + (index + 6)].Value = "FIRMA Y SELLO  DEL REPRESENTANTE LEGAL DE LA EMPRESA O EL QUE HAGA SUS VECES";
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Style.Font.Bold = true;
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Merge = true;
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Style.WrapText = true;
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Style.WrapText = true;
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C" + (index + 6) + ":F" + (index + 8)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
      


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
            workSheet.Column(1).Width = 5.57 + 0.71;
            workSheet.Column(2).Width = 15.00 + 0.71;
            workSheet.Column(3).Width = 45.43 + 0.71;
            workSheet.Column(4).Width = 8.86 + 0.71;

            workSheet.Column(5).Width = 9.43 + 0.71;
            workSheet.Column(6).Width = 9.00 + 0.71;
            workSheet.Column(7).Width = 10.86 + 0.71;
            workSheet.Column(8).Width = 9.29 + 0.71;
            workSheet.Column(9).Width = 12.29 + 0.71;
            workSheet.Column(10).Width = 17.86 + 0.71;
            workSheet.Column(11).Width = 13 + 0.71;
            workSheet.Column(12).Width = 9.86 + 0.71;


            workSheet.Column(13).Width = 11.14 + 0.71;
            workSheet.Column(14).Width = 13.00 + 0.71;
            workSheet.Column(15).Width = 16.14 + 0.71;
            workSheet.Column(16).Width = 15.29 + 0.71;
            workSheet.Column(17).Width = 15.14 + 0.71;
            workSheet.Column(18).Width = 15.00 + 0.71;
            workSheet.Column(19).Width = 13.57 + 0.71;
            workSheet.Column(20).Width = 9.29 + 0.71;

            workSheet.Column(21).Width = 8.14 + 0.71;
            workSheet.Column(22).Width = 22.14 + 0.71;

            workSheet.Row(1).Height = 74.25;
            workSheet.Row(2).Height = 15;
            workSheet.Row(3).Height = 36.00;
            workSheet.Row(6).Height = 103.50; 
            workSheet.Row(7).Height = 14.25;
            workSheet.Row(8).Height = 21.00;
            workSheet.Row(9).Height = 21.75;
            workSheet.Row(10).Height = 32.25;
            workSheet.Row(11).Height = 35.25;
            workSheet.Row(12).Height = 15.75;

            workSheet.Row(13).Height = 15.75;
            workSheet.Row(14).Height = 15.75;
            workSheet.Row(15).Height = 15.00;
            workSheet.Row(16).Height = 60.00;
            workSheet.Row(17).Height = 72.00;
         

        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {

            workSheet.Cells["A2:C2"].Merge = true;
            workSheet.Cells["A3:C3"].Merge = true;
            workSheet.Cells["A4:C4"].Merge = true;
            workSheet.Cells["D2:S4"].Merge = true;
            workSheet.Cells["A6:V6"].Merge = true;
            workSheet.Cells["A7:L7"].Merge = true;
            workSheet.Cells["C8:I8"].Merge = true;
            workSheet.Cells["M8:N8"].Merge = true;

            workSheet.Cells["A8:B8"].Merge = true;
            workSheet.Cells["C8:I8"].Merge = true;
            workSheet.Cells["K8:L8"].Merge = true;

            workSheet.Cells["A9:B9"].Merge = true;
            workSheet.Cells["C9:I9"].Merge = true;
            workSheet.Cells["K9:L9"].Merge = true;
            workSheet.Cells["M9:N9"].Merge = true;

            workSheet.Cells["A10:B10"].Merge = true;
            workSheet.Cells["C10:I10"].Merge = true;
            workSheet.Cells["K10:L10"].Merge = true;
            workSheet.Cells["M10:N10"].Merge = true;

            workSheet.Cells["A11:B11"].Merge = true;
            workSheet.Cells["C11:I11"].Merge = true;
            workSheet.Cells["K11:L11"].Merge = true;
            workSheet.Cells["M11:N11"].Merge = true;

            workSheet.Cells["A12:B12"].Merge = true;
            workSheet.Cells["D12:F12"].Merge = true;
            workSheet.Cells["G12:L12"].Merge = true;
            workSheet.Cells["A13:V13"].Merge = true;

            workSheet.Cells["A14:E14"].Merge = true;

            workSheet.Cells["F14:S14"].Merge = true;


            //DETALLE 

            workSheet.Cells["A15:A18"].Merge = true;
            workSheet.Cells["F15:S15"].Merge = true;
            workSheet.Cells["B16:B18"].Merge = true;

            workSheet.Cells["C16:F18"].Merge = true;
            workSheet.Cells["G16:H18"].Merge = true;

            workSheet.Cells["I16:I18"].Merge = true;
            workSheet.Cells["J16:J18"].Merge = true;
            workSheet.Cells["K16:K18"].Merge = true;
            workSheet.Cells["L16:M18"].Merge = true;
            workSheet.Cells["N16:N18"].Merge = true;
            workSheet.Cells["O16:O18"].Merge = true;
            workSheet.Cells["P16:P18"].Merge = true;
            workSheet.Cells["Q16:Q18"].Merge = true;
            workSheet.Cells["R16:R18"].Merge = true;
            workSheet.Cells["S16:S18"].Merge = true;
            workSheet.Cells["T14:U18"].Merge = true;
            workSheet.Cells["V14:V18"].Merge = true;

        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["D2:S4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A7:L7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A8:B8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A9:B9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A10:B10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A11:B11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A12:B12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D12:F12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G12:L12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["C8:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C9:I9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C10:I10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C11:I11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          

            workSheet.Cells["J8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K8:L8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K9:L9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K10:L10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K11:L11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M8:N8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M9:N9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M10:N10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M11:N11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O8:Q8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O9:Q9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O10:P10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O11:P11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["O11:P11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A14:E14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F14:S14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            //DETALLE

            workSheet.Cells["A15:A18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F15:S15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["B16:B18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C16:F18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G16:H18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I16:I18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J16:J18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K16:K18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L16:M18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N16:N18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O16:O18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P16:P18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q16:Q18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R16:R18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S16:S18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["T14:U18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["V14:V18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }

        private static void TextoNegrita(ExcelWorksheet workSheet)
        {

        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["D2:S4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
            workSheet.Cells["A7:L7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
            workSheet.Cells["A8:B8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["A9:B9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["A10:B10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["A11:B11"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["A12:B12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["D12:F12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));


            workSheet.Cells["J8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["J9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["J10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["J11"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            workSheet.Cells["M8:N8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["M9:N9"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["M10:N10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["M11:N11"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["A14:E14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));

            workSheet.Cells["F14:S14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));

            //DETALLE
            workSheet.Cells["A15:A18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["F15:S15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            workSheet.Cells["B16:B18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["C16:F18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));


            workSheet.Cells["G16:H18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["I16:I18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["J16:J18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["K16:K18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            workSheet.Cells["L16:N18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["O16:O18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["P16:P18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["Q16:Q18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["R16:R18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["S16:S18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["T14:U18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            workSheet.Cells["V14:V18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
        }

    }
}