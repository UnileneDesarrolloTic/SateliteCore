using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato9_Report
    {

        public static string Exportar(Image firma, Image logoUnilene, Coti_Formato_9_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización  essalud Apurimac");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(200, 50);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["A5"].Value = "DECLARACION JURADA DE COTIZACION";
                workSheet.Cells["A5"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A5"].Style.Font.Size = 16;
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A5"].Style.Font.Bold = true;

                workSheet.Cells["A6"].Value = "ADQUISICION MATERIAL MEDICO DELEGADO A COMPRA LOCAL";
                workSheet.Cells["A6"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A6"].Style.Font.Size = 14;
                workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A6"].Style.Font.Bold = true;

                workSheet.Cells["A8"].Value = "REQUERIMIENTO  NOTA N° 044-RR.MM-OPC-RAAP-ESSALUD-2021";
                workSheet.Cells["A8"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A8"].Style.Font.Size = 10;
                workSheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A8"].Style.Font.Bold = true;

                workSheet.Cells["A9"].Value = "NIT" +  cotizacion.Prov_nit;
                workSheet.Cells["A9"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A9"].Style.Font.Size = 10;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A9"].Style.Font.Bold = true;


                workSheet.Cells["N9"].Value = "Nro Cotizacion";
                workSheet.Cells["N9"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["N9"].Style.Font.Size = 14;
                workSheet.Cells["N9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N9"].Style.Font.Bold = true;

                workSheet.Cells["Q9"].Value = cotizacion.Prov_NroCotizacion+ "-2022 Unilene S.A.C";
                workSheet.Cells["Q9"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["Q9"].Style.Font.Size = 14;
                workSheet.Cells["Q9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q9"].Style.Font.Bold = true;



                workSheet.Cells["A10"].Value = "AREA REQUIRIENTE :" + cotizacion.Prov_AreaRequiriente;
                workSheet.Cells["A10"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A10"].Style.Font.Size = 14;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A10"].Style.Font.Bold = true;

                workSheet.Cells["N10"].Value = "Fecha: ";
                workSheet.Cells["N10"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["N10"].Style.Font.Size = 14;
                workSheet.Cells["N10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N10"].Style.Font.Bold = true;

                workSheet.Cells["Q10"].Value = cotizacion.Prov_Fecha;
                workSheet.Cells["Q10"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["Q10"].Style.Font.Size = 10;
                workSheet.Cells["Q10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q10"].Style.Font.Bold = true;
                workSheet.Cells["Q10"].Style.Numberformat.Format = "dd/MM/yyyy";


                workSheet.Cells["A11"].Value = "RESPONZABLE: " + cotizacion.Prov_Resposanble;
                workSheet.Cells["A11"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A11"].Style.Font.Size = 10;
                workSheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A11"].Style.Font.Bold = true;



                workSheet.Cells["A13"].Value = "Señores";
                workSheet.Cells["A13"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A13"].Style.Font.Size = 12;
                workSheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                


                workSheet.Cells["A14"].Value = "PROVEEDORES:   ";
                workSheet.Cells["A14"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A14"].Style.Font.Size = 12;
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
               

                workSheet.Cells["A15"].Value = "Presente: ";
                workSheet.Cells["A15"].Style.Font.Name = "Times New Roman";
                workSheet.Cells["A15"].Style.Font.Size = 12;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["A17"].Value = "SIRVANSE COTIZAR LOS SIGUIENTES ITEMS: (PRECIO COMPRA DIRECTA) 'BIENES'";
                workSheet.Cells["A17"].Style.Font.Name = "Calibri";
                workSheet.Cells["A17"].Style.Font.Size = 12;
                workSheet.Cells["A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                workSheet.Cells["O17"].Value = "Documentos Tecnicos del Postor";
                workSheet.Cells["O17"].Style.Font.Name = "Calibri";
                workSheet.Cells["O17"].Style.Font.Size = 12;
                workSheet.Cells["O17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O17"].Style.WrapText = true;

                workSheet.Cells["Q17"].Value = "Documentos Tecnicos del Postor";
                workSheet.Cells["Q17"].Style.Font.Name = "Calibri";
                workSheet.Cells["Q17"].Style.Font.Size = 12;
                workSheet.Cells["Q17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q17"].Style.WrapText = true;

                //DETALLE
                workSheet.Cells["A18"].Value = "ITEM";
                workSheet.Cells["A18"].Style.Font.Size = 10;
                workSheet.Cells["A18"].Style.Font.Name = "Calibri";
                workSheet.Cells["A18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A18"].Style.WrapText = true;
                workSheet.Cells["A18"].Style.Font.Bold = true;


                workSheet.Cells["B18"].Value = "SOLPE";
                workSheet.Cells["B18"].Style.Font.Size = 10;
                workSheet.Cells["B18"].Style.Font.Name = "Calibri";
                workSheet.Cells["B18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B18"].Style.WrapText = true;
                workSheet.Cells["B18"].Style.Font.Bold = true;

                workSheet.Cells["C18"].Value = "POS";
                workSheet.Cells["C18"].Style.Font.Size = 10;
                workSheet.Cells["C18"].Style.Font.Name = "Calibri";
                workSheet.Cells["C18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C18"].Style.WrapText = true;
                workSheet.Cells["C18"].Style.Font.Bold = true;

                workSheet.Cells["D18"].Value = "Codigo SAP";
                workSheet.Cells["D18"].Style.Font.Size = 10;
                workSheet.Cells["D18"].Style.Font.Name = "Calibri";
                workSheet.Cells["D18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D18"].Style.WrapText = true;
                workSheet.Cells["D18"].Style.Font.Bold = true;


                workSheet.Cells["E18"].Value = "DESCRIPCION ";
                workSheet.Cells["E18"].Style.Font.Size = 10;
                workSheet.Cells["E18"].Style.Font.Name = "Calibri";
                workSheet.Cells["E18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E18"].Style.WrapText = true;
                workSheet.Cells["E18"].Style.Font.Bold = true;

                workSheet.Cells["F18"].Value = "U.M";
                workSheet.Cells["F18"].Style.Font.Size = 10;
                workSheet.Cells["F18"].Style.Font.Name = "Calibri";
                workSheet.Cells["F18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F18"].Style.WrapText = true;
                workSheet.Cells["F18"].Style.Font.Bold = true;

                workSheet.Cells["G18"].Value = "Cantidad";
                workSheet.Cells["G18"].Style.Font.Size = 10;
                workSheet.Cells["G18"].Style.Font.Name = "Calibri";
                workSheet.Cells["G18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G18"].Style.WrapText = true;
                workSheet.Cells["G18"].Style.Font.Bold = true;

                workSheet.Cells["H18"].Value = "Precio Unit.";
                workSheet.Cells["H18"].Style.Font.Size = 10;
                workSheet.Cells["H18"].Style.Font.Name = "Calibri";
                workSheet.Cells["H18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H18"].Style.WrapText = true;
                workSheet.Cells["H18"].Style.Font.Bold = true;


                workSheet.Cells["I18"].Value = "Precio Total";
                workSheet.Cells["I18"].Style.Font.Size = 10;
                workSheet.Cells["I18"].Style.Font.Name = "Calibri";
                workSheet.Cells["I18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I18"].Style.WrapText = true;
                workSheet.Cells["I18"].Style.Font.Bold = true;

                workSheet.Cells["J18"].Value = "Marca del Producto";
                workSheet.Cells["J18"].Style.Font.Size = 10;
                workSheet.Cells["J18"].Style.Font.Name = "Calibri";
                workSheet.Cells["J18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J18"].Style.WrapText = true;
                workSheet.Cells["J18"].Style.Font.Bold = true;

                workSheet.Cells["K18"].Value = "Laboratorio";
                workSheet.Cells["K18"].Style.Font.Size = 10;
                workSheet.Cells["K18"].Style.Font.Name = "Calibri";
                workSheet.Cells["K18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K18"].Style.WrapText = true;
                workSheet.Cells["K18"].Style.Font.Bold = true;

                workSheet.Cells["L18"].Value = "Forma Presentacion";
                workSheet.Cells["L18"].Style.Font.Size = 10;
                workSheet.Cells["L18"].Style.Font.Name = "Calibri";
                workSheet.Cells["L18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L18"].Style.WrapText = true;
                workSheet.Cells["L18"].Style.Font.Bold = true;

                workSheet.Cells["M18"].Value = "Procedencia";
                workSheet.Cells["M18"].Style.Font.Size = 10;
                workSheet.Cells["M18"].Style.Font.Name = "Calibri";
                workSheet.Cells["M18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M18"].Style.WrapText = true;
                workSheet.Cells["M18"].Style.Font.Bold = true;


                workSheet.Cells["N18"].Value = "Plazo de Entrega en Dias Calendarios";
                workSheet.Cells["N18"].Style.Font.Size = 10;
                workSheet.Cells["N18"].Style.Font.Name = "Calibri";
                workSheet.Cells["N18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N18"].Style.WrapText = true;
                workSheet.Cells["N18"].Style.Font.Bold = true;


                workSheet.Cells["O18"].Value = "Resolucion de Autorización Sanitaria de Funcionamiento  de Establecimiento Farmacéutico o Constancia de Registro de Establecimiento Farmacéutico";
                workSheet.Cells["O18"].Style.Font.Size = 10;
                workSheet.Cells["O18"].Style.Font.Name = "Calibri";
                workSheet.Cells["O18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O18"].Style.WrapText = true;
                workSheet.Cells["O18"].Style.Font.Bold = true;

                workSheet.Cells["P18"].Value = "Certificado de Buenas Pracicas  de Almacenamiento(BPA)";
                workSheet.Cells["P18"].Style.Font.Size = 10;
                workSheet.Cells["P18"].Style.Font.Name = "Calibri";
                workSheet.Cells["P18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P18"].Style.WrapText = true;
                workSheet.Cells["P18"].Style.Font.Bold = true;


                workSheet.Cells["Q18"].Value = "Registro Sanitario o Certificado de Registro Sanitario";
                workSheet.Cells["Q18"].Style.Font.Size = 10;
                workSheet.Cells["Q18"].Style.Font.Name = "Calibri";
                workSheet.Cells["Q18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q18"].Style.WrapText = true;
                workSheet.Cells["Q18"].Style.Font.Bold = true;

                workSheet.Cells["R18"].Value = "Certificado de Buenas Practicas de Manufactura (BPM)";
                workSheet.Cells["R18"].Style.Font.Size = 10;
                workSheet.Cells["R18"].Style.Font.Name = "Calibri";
                workSheet.Cells["R18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R18"].Style.WrapText = true;
                workSheet.Cells["R18"].Style.Font.Bold = true;

                workSheet.Cells["S18"].Value = "Certificado de Análisis del Producto Terminado (Protocolo de Análisis)";
                workSheet.Cells["S18"].Style.Font.Size = 10;
                workSheet.Cells["S18"].Style.Font.Name = "Calibri";
                workSheet.Cells["S18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S18"].Style.WrapText = true;
                workSheet.Cells["S18"].Style.Font.Bold = true;


                workSheet.Cells["T18"].Value = "Vigencia Mínima del Producto  a la entrega   (=>18 meses)";
                workSheet.Cells["T18"].Style.Font.Size = 10;
                workSheet.Cells["T18"].Style.Font.Name = "Calibri";
                workSheet.Cells["T18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T18"].Style.WrapText = true;
                workSheet.Cells["T18"].Style.Font.Bold = true;


                workSheet.Cells["U18"].Value = "declaro cumplir al 100% con las Especificaciones Tecnicas (si-no)";
                workSheet.Cells["U18"].Style.Font.Size = 10;
                workSheet.Cells["U18"].Style.Font.Name = "Calibri";
                workSheet.Cells["U18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["U18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["U18"].Style.WrapText = true;
                workSheet.Cells["U18"].Style.Font.Bold = true;

                workSheet.Cells["V18"].Value = "cuento con Registro Nacional de Proveedores (RNP) Vigente (si-no)";
                workSheet.Cells["V18"].Style.Font.Size = 10;
                workSheet.Cells["V18"].Style.Font.Name = "Calibri";
                workSheet.Cells["V18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["V18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["V18"].Style.WrapText = true;
                workSheet.Cells["V18"].Style.Font.Bold = true;

                workSheet.Cells["W18"].Value = "Observaciones";
                workSheet.Cells["W18"].Style.Font.Size = 10;
                workSheet.Cells["W18"].Style.Font.Name = "Calibri";
                workSheet.Cells["W18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["W18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["W18"].Style.WrapText = true;
                workSheet.Cells["W18"].Style.Font.Bold = true;

                int index = 19;
                foreach (Coti_Formato9_Detalle item in cotizacion.Detalle)
                {

                    workSheet.Row(index).Height = 28.5;
                    workSheet.Cells["A" + index].Value = index - 18;
                    workSheet.Cells["A" + index].Style.Font.Size = 10;
                    workSheet.Cells["A" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.WrapText = true;
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";

                    workSheet.Cells["B" + index].Value = item.Solpe;
                    workSheet.Cells["B" + index].Style.Font.Size = 9;
                    workSheet.Cells["B" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;

                    workSheet.Cells["C" + index].Value = item.Pos;
                    workSheet.Cells["C" + index].Style.Font.Size = 9;
                    workSheet.Cells["C" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;


                    workSheet.Cells["D" + index].Value = item.CodigoSAP;
                    workSheet.Cells["D" + index].Style.Font.Size = 9;
                    workSheet.Cells["D" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;


                    workSheet.Cells["E" + index].Value = item.Descripcion;
                    workSheet.Cells["E" + index].Style.Font.Size = 9;
                    workSheet.Cells["E" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;

                    workSheet.Cells["F" + index].Value = item.Um;
                    workSheet.Cells["F" + index].Style.Font.Size = 9;
                    workSheet.Cells["F" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;

                    workSheet.Cells["G" + index].Value = item.Cantidad;
                    workSheet.Cells["G" + index].Style.Font.Size = 9;
                    workSheet.Cells["G" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;
                    workSheet.Cells["G" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["H" + index].Value = item.PreUnitario;
                    workSheet.Cells["H" + index].Style.Font.Size = 9;
                    workSheet.Cells["H" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;
                    workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["I" + index].Value = item.PreTotal;
                    workSheet.Cells["I" + index].Style.Font.Size = 9;
                    workSheet.Cells["I" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;
                    workSheet.Cells["I" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["J" + index].Value = item.Marca;
                    workSheet.Cells["J" + index].Style.Font.Size = 9;
                    workSheet.Cells["J" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;

                    workSheet.Cells["K" + index].Value = item.Laboratorio;
                    workSheet.Cells["K" + index].Style.Font.Size = 9;
                    workSheet.Cells["K" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;

                    workSheet.Cells["L" + index].Value = item.Presentacion;
                    workSheet.Cells["L" + index].Style.Font.Size = 9;
                    workSheet.Cells["L" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;

                    workSheet.Cells["M" + index].Value = item.Procedencia;
                    workSheet.Cells["M" + index].Style.Font.Size = 9;
                    workSheet.Cells["M" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    workSheet.Cells["N" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["N" + index].Style.Font.Size = 9;
                    workSheet.Cells["N" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;

                    workSheet.Cells["O" + index].Value = item.ResolAutSanitaria;
                    workSheet.Cells["O" + index].Style.Font.Size = 9;
                    workSheet.Cells["O" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;

                    workSheet.Cells["P" + index].Value = item.CertBPA;
                    workSheet.Cells["P" + index].Style.Font.Size = 9;
                    workSheet.Cells["P" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["P" + index].Style.WrapText = true;

                    workSheet.Cells["Q" + index].Value = item.RegSanitario;
                    workSheet.Cells["Q" + index].Style.Font.Size = 9;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;


                    workSheet.Cells["R" + index].Value = item.CertBPM;
                    workSheet.Cells["R" + index].Style.Font.Size = 9;
                    workSheet.Cells["R" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["R" + index].Style.WrapText = true;

                    workSheet.Cells["S" + index].Value = item.ProtocoloAnalisis;
                    workSheet.Cells["S" + index].Style.Font.Size = 9;
                    workSheet.Cells["S" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["S" + index].Style.WrapText = true;

                    workSheet.Cells["T" + index].Value = item.VigMinimaProducto;
                    workSheet.Cells["T" + index].Style.Font.Size = 9;
                    workSheet.Cells["T" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["T" + index].Style.WrapText = true;

                    workSheet.Cells["U" + index].Value = item.EspTecnicas;
                    workSheet.Cells["U" + index].Style.Font.Size = 9;
                    workSheet.Cells["U" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["U" + index].Style.WrapText = true;


                    workSheet.Cells["V" + index].Value = item.RegNacionalProveedores;
                    workSheet.Cells["V" + index].Style.Font.Size = 9;
                    workSheet.Cells["V" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["V" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["V" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["V" + index].Style.WrapText = true;


                    workSheet.Cells["W" + index].Value = item.RegNacionalProveedores;
                    workSheet.Cells["W" + index].Style.Font.Size = 9;
                    workSheet.Cells["W" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["W" + index].Style.WrapText = true;


                    index++;


                }


                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);
                PintarCeldas(workSheet);

                workSheet.Cells["H" + index].Value = "TOTAL S/";
                workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H" + index].Style.Font.Bold = true;
                workSheet.Cells["H" + index].Style.Font.Size = 10;
                workSheet.Cells["H" + index].Style.Font.Name = "Calibri";

                workSheet.Cells["I" + index].Value = cotizacion.Prov_valorTotal;
                workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I" + index].Style.Font.Bold = true;
                workSheet.Cells["I" + index].Style.Font.Size = 10;
                workSheet.Cells["I" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["I" + index].Style.Numberformat.Format = "#,##0.00";

                index++;
                workSheet.Cells["B" + index + ":V" + index].Merge = true;
                workSheet.Cells["B" + index + ":V" + index].Value = "1.- El precio deberá ser el mas competitivo del mercado";
               
                workSheet.Cells["B" + index + ":V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Name = "Calibri";

                index++;
                
                workSheet.Cells["B" + index + ":V" + index].Merge = true;
                workSheet.Cells["B" + index + ":V" + index].Value = "2.- Deberan cotizar solo los que cumplen con ' todas las condiciones de las especificaciones técnicas', las mismas que se han remitido adjunto a la presente y que; con el envio de la presente Declaración Jurada certifica que son de su total conociemnto.";
                
                workSheet.Cells["B" + index + ":V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Name = "Calibri";

                index++;
                workSheet.Cells["B" + index + ":V" + index].Merge = true;
                workSheet.Cells["B" + index + ":V" + index].Value = "3.- Conservar el mismo formato ";
               
                workSheet.Cells["B" + index + ":V" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Bold = true;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":V" + index].Style.Font.Name = "Calibri";

                index++;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "PROVEEDOR";
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.WrapText = true;


                workSheet.Cells["E" + index + ":H" + index].Merge = true;
                workSheet.Cells["E" + index + ":H" + index].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["E" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Size = 10;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["E" + index + ":H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "RUC";
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.WrapText = true;


                workSheet.Cells["E" + index + ":H" + index].Merge = true;
                workSheet.Cells["E" + index + ":H" + index].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["E" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Size = 10;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["E" + index + ":H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                index++;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "TELEFONO-CELULAR";
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.WrapText = true;


                workSheet.Cells["E" + index + ":H" + index].Merge = true;
                workSheet.Cells["E" + index + ":H" + index].Value = cotizacion.Prov_telefono+" - " + cotizacion.Prov_movil;
                workSheet.Cells["E" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Size = 10;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["E" + index + ":H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index++;
                workSheet.Cells["B" + index + ":D" + index].Merge = true;
                workSheet.Cells["B" + index + ":D" + index].Value = "CORREO ELECTRONICO";
                workSheet.Cells["B" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Size = 10;
                workSheet.Cells["B" + index + ":D" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["B" + index + ":D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + index + ":D" + index].Style.WrapText = true;


                workSheet.Cells["E" + index + ":H" + index].Merge = true;
                workSheet.Cells["E" + index + ":H" + index].Value = cotizacion.Prov_Email;
                workSheet.Cells["E" + index + ":H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Size = 10;
                workSheet.Cells["E" + index + ":H" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["E" + index + ":H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                index = index + 2;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "Suministrar información a:";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Bold =true;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "ESSALUD - RED ASISTENCIAL APURIMAC";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "URB. SAN FRANCISCO LOTE A, PARTE INTEGRANTE DEL LOTE 61 Y 61-A SECTOR PATIBAMBA ABANCAY-APURÍMAC";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "UNIDAD DE ADQUISICIONES INGENIERIA HOSPITALARIA Y SERVICIOS";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                index++;

                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "Informes (Teléfono/email):";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "083-595555 ANEXO 1023";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "email: programacion.apurimac@essalud.gob.pe";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "email (alterntivo): compras.apurimac@gmail.com";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "Observaciones: ";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":P" + index].Merge = true;
                workSheet.Cells["A" + index + ":P" + index].Value = "SIRVACE COTIZAR EN NUEVOS SOLES, INCLUIR IGV y TODO COSTO.";
                workSheet.Cells["A" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":P" + index].Style.Font.Name = "Calibri";
                workSheet.Cells["A" + index + ":P" + index].Style.WrapText = true;

                ExcelPicture firmaCatizacion = workSheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(index-16, -5, 8, 60);
                firmaCatizacion.SetSize(265, 120);

                workSheet.View.ZoomScale = 85;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

            }
            return reporte;

        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Column(1).Width = 5 + 0.71;
            workSheet.Column(2).Width = 8.57 + 0.71;
            workSheet.Column(3).Width = 4.29 + 0.71;
            workSheet.Column(4).Width = 12.86 + 0.71;
            workSheet.Column(5).Width = 33.71 + 0.71;
            workSheet.Column(6).Width = 5.86 + 0.71;
            workSheet.Column(7).Width = 7.86 + 0.71;
            workSheet.Column(8).Width = 8.14 + 0.71;

            workSheet.Column(9).Width = 9.43 + 0.71;
            workSheet.Column(10).Width = 8.29 + 0.71;
            workSheet.Column(11).Width = 10 + 0.71;
            workSheet.Column(12).Width = 11.43 + 0.71;
            workSheet.Column(13).Width = 10.71 + 0.71;
            workSheet.Column(14).Width = 11 + 0.71;
            workSheet.Column(15).Width = 19.14 + 0.71;

            workSheet.Column(16).Width = 11.57 + 0.71;
            workSheet.Column(17).Width = 11.57 + 0.71;
            workSheet.Column(18).Width = 11.57 + 0.71;
            workSheet.Column(19).Width = 11.57 + 0.71;
            workSheet.Column(20).Width = 11.57 + 0.71;
            workSheet.Column(21).Width = 14.14 + 0.71;
            workSheet.Column(22).Width = 15.71 + 0.71;
            workSheet.Column(23).Width = 12.29 + 0.71;

            workSheet.Row(1).Height = 18.75;
            workSheet.Row(2).Height = 20.25;
            workSheet.Row(3).Height = 20.25;
            workSheet.Row(4).Height = 20.25; 
            workSheet.Row(5).Height = 22.5; 
            workSheet.Row(6).Height = 22.5;
            workSheet.Row(7).Height = 18.75;
            workSheet.Row(8).Height = 18.75;


            workSheet.Row(9).Height = 18.75;
            workSheet.Row(10).Height = 18.75;
            workSheet.Row(11).Height = 15.75;
            workSheet.Row(12).Height = 15.75;
            workSheet.Row(13).Height = 15.75;
            workSheet.Row(14).Height = 15.75;
            workSheet.Row(15).Height = 15.75;
            workSheet.Row(16).Height = 10.5;
            workSheet.Row(17).Height = 20.25;
            workSheet.Row(18).Height = 114.00;
        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {

            workSheet.Cells["A5:J5"].Merge = true;
            workSheet.Cells["A6:J6"].Merge = true;
            workSheet.Cells["A8:E8"].Merge = true;
            workSheet.Cells["A9:E9"].Merge = true;
            workSheet.Cells["N9:P9"].Merge = true;
            workSheet.Cells["Q9:S9"].Merge = true;

            workSheet.Cells["A10:E10"].Merge = true;
            workSheet.Cells["A11:E11"].Merge = true;

            workSheet.Cells["N10:P10"].Merge = true;
            workSheet.Cells["Q10:S10"].Merge = true;

            workSheet.Cells["A17:N17"].Merge = true;

            workSheet.Cells["O17:P17"].Merge = true;
            workSheet.Cells["Q17:T17"].Merge = true;

        }
        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["N9:P9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q9:S9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["N10:P10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q10:S10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["O17:P17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q17:T17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["J18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["M18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["O18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["Q18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["T18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["U18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["V18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["W18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["O17:P17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B7DEE8"));
            workSheet.Cells["Q17:T17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FABF8F"));

            workSheet.Cells["A18:N18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

            workSheet.Cells["O18:P18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B7DEE8"));
            workSheet.Cells["Q18:T18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FABF8F"));

            workSheet.Cells["U18:W18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
        }
    }
 }
