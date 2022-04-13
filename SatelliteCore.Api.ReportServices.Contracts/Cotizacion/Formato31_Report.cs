using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato31_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato31_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Tumbes");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(150, 50);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);



                workSheet.Cells["A2"].Value = "SOLICITUD DE COTIZACIÓN Y DECLARACION JURADA DE CUMPLIMIENTO DE ESPECIFICACIONES TECNICAS ";
                workSheet.Cells["A2"].Style.Font.Size = 16;
                workSheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A2"].Style.Font.Bold = true;


                workSheet.Cells["A5"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A5"].Style.Font.Size = 10;
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A5"].Style.Font.Bold = true;

                workSheet.Cells["K5"].Value = "DATOS DEL CONTACTO";
                workSheet.Cells["K5"].Style.Font.Size = 10;
                workSheet.Cells["K5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K5"].Style.Font.Bold = true;


                workSheet.Cells["A6"].Value = "RAZON SOCIAL";
                workSheet.Cells["A6"].Style.Font.Size = 10;
                workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A6"].Style.Font.Bold = true;

                workSheet.Cells["E6"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["E6"].Style.Font.Size = 10;
                workSheet.Cells["E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E6"].Style.Font.Bold = true;

                workSheet.Cells["K6"].Value = "NOMBRE";
                workSheet.Cells["K6"].Style.Font.Size = 10;
                workSheet.Cells["K6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K6"].Style.Font.Bold = true;


                workSheet.Cells["L6"].Value = cotizacion.Prov_Nombre;
                workSheet.Cells["L6"].Style.Font.Size = 10;
                workSheet.Cells["L6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L6"].Style.Font.Bold = true;


                workSheet.Cells["A7"].Value = "DIRECCION";
                workSheet.Cells["A7"].Style.Font.Size = 10;
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A7"].Style.Font.Bold = true;

                workSheet.Cells["E7"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["E7"].Style.Font.Size = 10;
                workSheet.Cells["E7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E7"].Style.Font.Bold = true;

                workSheet.Cells["K7"].Value = "CARGO";
                workSheet.Cells["K7"].Style.Font.Size = 10;
                workSheet.Cells["K7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K7"].Style.Font.Bold = true;


                workSheet.Cells["L7"].Value = cotizacion.Prov_Cargo;
                workSheet.Cells["L7"].Style.Font.Size = 10;
                workSheet.Cells["L7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L7"].Style.Font.Bold = true;


                workSheet.Cells["A8"].Value = "TELEFONO (S)";
                workSheet.Cells["A8"].Style.Font.Size = 10;
                workSheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A8"].Style.Font.Bold = true;

                workSheet.Cells["E8"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["E8"].Style.Font.Size = 10;
                workSheet.Cells["E8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E8"].Style.Font.Bold = true;


                workSheet.Cells["K8"].Value = "MOVIL (Indicar RPM ó RPC)";
                workSheet.Cells["K8"].Style.Font.Size = 10;
                workSheet.Cells["K8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K8"].Style.Font.Bold = true;


                workSheet.Cells["N8"].Value = cotizacion.Prov_Movil;
                workSheet.Cells["N8"].Style.Font.Size = 10;
                workSheet.Cells["N8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N8"].Style.Font.Bold = true;


                workSheet.Cells["A9"].Value = "EMAIL EMPRESA:";
                workSheet.Cells["A9"].Style.Font.Size = 10;
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A9"].Style.Font.Bold = true;

                workSheet.Cells["E9"].Value = cotizacion.Prov_Email;
                workSheet.Cells["E9"].Style.Font.Size = 10;
                workSheet.Cells["E9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E9"].Style.Font.Bold = true;


                workSheet.Cells["K9"].Value = "E MAIL DEL CONTACTO";
                workSheet.Cells["K9"].Style.Font.Size = 10;
                workSheet.Cells["K9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K9"].Style.Font.Bold = true;


                workSheet.Cells["N9"].Value = cotizacion.Prov_EmailContacto;
                workSheet.Cells["N9"].Style.Font.Size = 10;
                workSheet.Cells["N9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N9"].Style.Font.Bold = true;




                workSheet.Cells["A10"].Value = "FECHA OFERTA:";
                workSheet.Cells["A10"].Style.Font.Size = 10;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A10"].Style.Font.Bold = true;

                workSheet.Cells["E10"].Value = cotizacion.Prov_FeOferta;
                workSheet.Cells["E10"].Style.Font.Size = 10;
                workSheet.Cells["E10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E10"].Style.Font.Bold = true;
                workSheet.Cells["E10"].Style.Numberformat.Format = "dd/MM/yyyy";


                workSheet.Cells["F10"].Value = "RUC:";
                workSheet.Cells["F10"].Style.Font.Size = 10;
                workSheet.Cells["F10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F10"].Style.Font.Bold = true;

                workSheet.Cells["G10"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["G10"].Style.Font.Size = 10;
                workSheet.Cells["G10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G10"].Style.Font.Bold = true;


                workSheet.Cells["K10"].Value = "PLAZO DE OFERTA";
                workSheet.Cells["K10"].Style.Font.Size = 10;
                workSheet.Cells["K10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K10"].Style.Font.Bold = true;


                workSheet.Cells["N10"].Value = cotizacion.Prov_PlazoOferta;
                workSheet.Cells["N10"].Style.Font.Size = 10;
                workSheet.Cells["N10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N10"].Style.Font.Bold = true;



                workSheet.Cells["A12"].Value = "ADQUISICIÓN DE MATERIAL MEDICO DELEGADO  - PARA LA RED ASISTENCIAL TUMBES";
                workSheet.Cells["A12"].Style.Font.Size = 16;
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A12"].Style.Font.Bold = true;

                workSheet.Cells["A13"].Value = "SR. PROVEEDOR:    Por medio del presente solicitamos se sirva llenar la solicitud de cotización para precios referenciales de lo siguiente:" +
                    " (en concordancia al Artículo 12° del Reglamento de la Ley de Contrataciones del Estado; es decir que su valor referencial deberá ser calculado a todo costo, incluyendo todos los tributos, seguros, transportes, inspecciones, pruebas y, de ser el caso, los costos laborales respectivos" +
                    " conforme a la legislación vigente, así como cualquier otro concepto que pueda incidir sobre el costo de los bienes)";
                workSheet.Cells["A13"].Style.Font.Size = 10;
                workSheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A13"].Style.WrapText = true;


                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                BordesCeldas(workSheet);
                PintarCeldas(workSheet);


                //DETALLE

                workSheet.Cells["A14"].Value = "N°";
                workSheet.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A14"].Style.Font.Size = 7;
                workSheet.Cells["A14"].Style.WrapText = true;
                workSheet.Cells["A14"].Style.Font.Bold = true;

                workSheet.Cells["B14"].Value = "Solicitud de pedido Delegado";
                workSheet.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B14"].Style.Font.Size = 7;
                workSheet.Cells["B14"].Style.WrapText = true;
                workSheet.Cells["B14"].Style.Font.Bold = true;
                workSheet.Cells["B14"].Style.TextRotation = 90;

                workSheet.Cells["C14"].Value = "Posición";
                workSheet.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C14"].Style.Font.Size = 7;
                workSheet.Cells["C14"].Style.WrapText = true;
                workSheet.Cells["C14"].Style.Font.Bold = true;
                workSheet.Cells["C14"].Style.TextRotation = 90;


                workSheet.Cells["D14"].Value = "DESCRIPCIÓN DEL ITEM";
                workSheet.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D14"].Style.Font.Size = 7;
                workSheet.Cells["D14"].Style.WrapText = true;
                workSheet.Cells["D14"].Style.Font.Bold = true;

                workSheet.Cells["D16"].Value = "CODIGO SAP";
                workSheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D16"].Style.Font.Size = 7;
                workSheet.Cells["D16"].Style.WrapText = true;
                workSheet.Cells["D16"].Style.Font.Bold = true;

                workSheet.Cells["E16"].Value = "PRODUCTO (ITEM)";
                workSheet.Cells["E16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E16"].Style.Font.Size = 9;
                workSheet.Cells["E16"].Style.WrapText = true;
                workSheet.Cells["E16"].Style.Font.Bold = true;


                workSheet.Cells["F16"].Value = "Cantidad Total";
                workSheet.Cells["F16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F16"].Style.Font.Size = 9;
                workSheet.Cells["F16"].Style.WrapText = true;



                workSheet.Cells["G14"].Value = "UM";
                workSheet.Cells["G14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G14"].Style.Font.Size = 9;
                workSheet.Cells["G14"].Style.WrapText = true;



                workSheet.Cells["H14"].Value = "Precio Unitario S/.";
                workSheet.Cells["H14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H14"].Style.Font.Size = 9;
                workSheet.Cells["H14"].Style.WrapText = true;
                


                workSheet.Cells["I14"].Value = "Precio Total S/.";
                workSheet.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I14"].Style.Font.Size = 9;
                workSheet.Cells["I14"].Style.WrapText = true;

                workSheet.Cells["J14"].Value = "Marca y/o laboratorio fabricante";
                workSheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J14"].Style.Font.Size = 9;
                workSheet.Cells["J14"].Style.WrapText = true;

                workSheet.Cells["K14"].Value = "Procedencia";
                workSheet.Cells["K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K14"].Style.Font.Size = 9;
                workSheet.Cells["K14"].Style.WrapText = true;
                workSheet.Cells["K14"].Style.TextRotation = 90;
                workSheet.Cells["K14"].Style.Font.Bold = true;

                workSheet.Cells["L14"].Value = "Forma de Presentación del Producto";
                workSheet.Cells["L14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L14"].Style.Font.Size = 9;
                workSheet.Cells["L14"].Style.WrapText = true;
                workSheet.Cells["L14"].Style.TextRotation = 90;
                workSheet.Cells["L14"].Style.Font.Bold = true;


                workSheet.Cells["M14"].Value = "Plazo de Entrega total del Producto";
                workSheet.Cells["M14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M14"].Style.Font.Size = 9;
                workSheet.Cells["M14"].Style.WrapText = true;
                workSheet.Cells["M14"].Style.TextRotation = 90;
                workSheet.Cells["M14"].Style.Font.Bold = true;


                workSheet.Cells["N14"].Value = "DECLARACION JURADA DE CUMPLIMIENTO DE EE.TT. MINIMAS";
                workSheet.Cells["N14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N14"].Style.Font.Size = 8;
                workSheet.Cells["N14"].Style.WrapText = true;
                workSheet.Cells["N14"].Style.Font.Bold = true;



                workSheet.Cells["N18"].Value = "La vigencia mínima de los materiales deberá ser de 18 meses. (salvo excepción) (Si o No)";
                workSheet.Cells["N18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N18"].Style.Font.Size = 8;
                workSheet.Cells["N18"].Style.WrapText = true;
                workSheet.Cells["N18"].Style.TextRotation = 90;

                workSheet.Cells["O18"].Value = "Cumple con Protocolo de Análisis y Registro Sanitario (traducido al castellano) Para recepción conforme en Almacén (sí o No)";
                workSheet.Cells["O18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O18"].Style.Font.Size = 8;
                workSheet.Cells["O18"].Style.WrapText = true;
                workSheet.Cells["O18"].Style.TextRotation = 90;

                workSheet.Cells["P18"].Value = "Cumple al 100% con la denominación, descripción y especificaciones técnicas del ítem (Si o No) si es No, indicar el detalle para evaluación";
                workSheet.Cells["P18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P18"].Style.Font.Size = 8;
                workSheet.Cells["P18"].Style.WrapText = true;
                workSheet.Cells["P18"].Style.TextRotation = 90;

                workSheet.Cells["Q18"].Value = "Capacidad de Atención al 100% de la Cantdidad Total solicitada (Si o No)";
                workSheet.Cells["Q18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q18"].Style.Font.Size = 8;
                workSheet.Cells["Q18"].Style.WrapText = true;
                workSheet.Cells["Q18"].Style.TextRotation = 90;


                workSheet.Cells["R18"].Value = "Cumple con los Plazos de Entrega establecidos en su cotización (Si  o No)";
                workSheet.Cells["R18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R18"].Style.Font.Size = 8;
                workSheet.Cells["R18"].Style.WrapText = true;
                workSheet.Cells["R18"].Style.TextRotation = 90;

              

                workSheet.Cells["S18"].Value = "Certificado de Buenas Prácticas de Almacenamiento igente (Si o No)";
                workSheet.Cells["S18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S18"].Style.Font.Size = 8;
                workSheet.Cells["S18"].Style.WrapText = true;
                workSheet.Cells["S18"].Style.TextRotation = 90;

                workSheet.Cells["T18"].Value = "REGISTRO SANITARIO VIGENTE (SI O NO)";
                workSheet.Cells["T18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T18"].Style.Font.Size = 8;
                workSheet.Cells["T18"].Style.WrapText = true;
                workSheet.Cells["T18"].Style.TextRotation = 90;

                workSheet.Cells["U18"].Value = "Cumple con las normas de conservación y transporte del producto y se compromete a canjear de haber alguna observación a la recepción del producto  (Sí o No)";
                workSheet.Cells["U18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["U18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["U18"].Style.Font.Size = 8;
                workSheet.Cells["U18"].Style.WrapText = true;
                workSheet.Cells["U18"].Style.TextRotation = 90;

                int index=21;

                foreach (Coti_Formato31_Detalle item in cotizacion.Detalle)
                {

                    workSheet.Row(index).Height = 16.5;
                    workSheet.Cells["A" + index].Value = index - 20;
                    workSheet.Cells["A" + index].Style.Font.Size = 9;
                    workSheet.Cells["A" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["A" + index].Style.WrapText = true;

                    workSheet.Cells["B" + index].Value = item.SoliPediDelegado;
                    workSheet.Cells["B" + index].Style.Font.Size = 9;
                    workSheet.Cells["B" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;

                    workSheet.Cells["C" + index].Value = item.Posicion;
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


                    workSheet.Cells["E" + index].Value = item.Productos;
                    workSheet.Cells["E" + index].Style.Font.Size = 9;
                    workSheet.Cells["E" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;

                    workSheet.Cells["F" + index].Value = item.Cantidad;
                    workSheet.Cells["F" + index].Style.Font.Size = 9;
                    workSheet.Cells["F" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;

                    workSheet.Cells["G" + index].Value = item.Um;
                    workSheet.Cells["G" + index].Style.Font.Size = 9;
                    workSheet.Cells["G" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;


                    workSheet.Cells["H" + index].Value = item.PreUnitario;
                    workSheet.Cells["H" + index].Style.Font.Size = 9;
                    workSheet.Cells["H" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;


                    workSheet.Cells["I" + index].Value = item.PreTotal;
                    workSheet.Cells["I" + index].Style.Font.Size = 9;
                    workSheet.Cells["I" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Cells["J" + index].Value = item.MarcaLaboratorio;
                    workSheet.Cells["J" + index].Style.Font.Size = 9;
                    workSheet.Cells["J" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;

                    workSheet.Cells["K" + index].Value = item.Procedencia;
                    workSheet.Cells["K" + index].Style.Font.Size = 9;
                    workSheet.Cells["K" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;

                    workSheet.Cells["L" + index].Value = item.ForPresentProducto;
                    workSheet.Cells["L" + index].Style.Font.Size = 9;
                    workSheet.Cells["L" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;

                    workSheet.Cells["M" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["M" + index].Style.Font.Size = 9;
                    workSheet.Cells["M" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    workSheet.Cells["N" + index].Value = item.VigMinima;
                    workSheet.Cells["N" + index].Style.Font.Size = 9;
                    workSheet.Cells["N" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;

                    workSheet.Cells["O" + index].Value = item.CumpAnalisis;
                    workSheet.Cells["O" + index].Style.Font.Size = 9;
                    workSheet.Cells["O" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;

                    workSheet.Cells["P" + index].Value = item.CumpDemoninacion;
                    workSheet.Cells["P" + index].Style.Font.Size = 9;
                    workSheet.Cells["P" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["P" + index].Style.WrapText = true;

                    workSheet.Cells["Q" + index].Value = item.CapAtencion;
                    workSheet.Cells["Q" + index].Style.Font.Size = 9;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;

                    workSheet.Cells["R" + index].Value = item.CumplPlazo;
                    workSheet.Cells["R" + index].Style.Font.Size = 9;
                    workSheet.Cells["R" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["R" + index].Style.WrapText = true;

                    workSheet.Cells["S" + index].Value = item.CertBPractica;
                    workSheet.Cells["S" + index].Style.Font.Size = 9;
                    workSheet.Cells["S" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["S" + index].Style.WrapText = true;

                    workSheet.Cells["T" + index].Value = item.RegiSanitario;
                    workSheet.Cells["T" + index].Style.Font.Size = 9;
                    workSheet.Cells["T" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["T" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["T" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["T" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["T" + index].Style.WrapText = true;

                    workSheet.Cells["U" + index].Value = item.CumpNormas;
                    workSheet.Cells["U" + index].Style.Font.Size = 9;
                    workSheet.Cells["U" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["U" + index].Style.WrapText = true;


                    index++;


                }

                workSheet.Row(index).Height = 40.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "NOTA :  La vigencia mínima de los dispositivos o medicinas deberá ser igual o mayor a dieciocho (18) meses " +
                    "al momento de sus fecha de entrega en nuestros almacenes, no se acepta materiales ni medicinas" +
                    " con próximo vencimiento. Solo se aceptará con carta de compromiso de canje.";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 14;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Bold = true;

                index++;
                workSheet.Row(index).Height = 16.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "CONSIDERACIONES GENERALES:";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.UnderLine = true;

                index++;
                workSheet.Row(index).Height = 16.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "1. El precio de Mercado será a todo costo, es decir, deberá incluir todos los tributos (incluido el I.G.V.), seguros, transportes, inspecciones, pruebas de control de calidad, de ser el caso los costos laborales respectivos conforme a la legislación vigente:";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 16.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "2. El Valor Referencial Total debe ser expresado como máximo en 2 (DOS) decimales, el precio unitario se podrá expresar hasta en 3 decimales.";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";

                index++;
                workSheet.Row(index).Height = 16.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "3. La compra se realizará considerando el 100% de la cantidad total consignada en el cuadro de requerimientos.";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";



                index++;
                workSheet.Row(index).Height = 16.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "En calidad de Proveedor, luego de haber examinado los documentos del proceso de la referencia proporcionados por EsSalud - Red Asistencial Tumbes, y conocer todas las condiciones existentes, el suscrito ofrece entregar de acuerdo a lo solicitado";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";


                index++;
                workSheet.Row(index).Height = 16.5;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Value = "En ese sentido, me comprometo a entregar el bien con las características, en la forma, plazo de entrega especificados en la presente.";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";




                workSheet.Cells["H" + (index+4) + ":K" + (index + 4)].Merge = true;
                workSheet.Cells["H" + (index + 4) + ":K" + (index + 4)].Value = "Firma del Representante de la Empresa";
                workSheet.Cells["H" + (index + 4) + ":K" + (index + 4)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H" + (index + 4) + ":K" + (index + 4)].Style.WrapText = true;
                workSheet.Cells["H" + (index + 4) + ":K" + (index + 4)].Style.Font.Size = 8;
                workSheet.Cells["H" + (index + 4) + ":K" + (index + 4)].Style.Font.Name = "Arial";



                workSheet.Cells["N" + (index + 4) + ":U" + (index + 4)].Merge = true;
                workSheet.Cells["N" + (index + 4) + ":U" + (index + 4)].Value = "Lugar …..…………….fecha, ….....de…….………….....del 2019.";
                workSheet.Cells["N" + (index + 4) + ":U" + (index + 4)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N" + (index + 4) + ":U" + (index + 4)].Style.WrapText = true;
                workSheet.Cells["N" + (index + 4) + ":U" + (index + 4)].Style.Font.Size = 8;
                workSheet.Cells["N" + (index + 4) + ":U" + (index + 4)].Style.Font.Name = "Arial";


                workSheet.View.ZoomScale = 85;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);


                return reporte;
            }
        }
        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet workSheet)
        {

            workSheet.Column(1).Width = 3.29 + 0.71;
            workSheet.Column(2).Width = 10.43 + 0.71;
            workSheet.Column(3).Width = 12.43  + 0.71;
            workSheet.Column(4).Width = 9.57 + 0.71;
            workSheet.Column(5).Width = 55.14 + 0.71;
            workSheet.Column(6).Width = 7.43 + 0.71;
            workSheet.Column(7).Width = 5.43  + 0.71;
            workSheet.Column(8).Width = 8.71  + 0.71;
            workSheet.Column(9).Width = 9.86  + 0.71;
            workSheet.Column(10).Width = 9.14 + 0.71;
            workSheet.Column(11).Width = 8.71 + 0.71; 
            workSheet.Column(12).Width = 7.86 + 0.71;
            workSheet.Column(13).Width = 10 + 0.71;
            workSheet.Column(14).Width = 7.29 + 0.71;
            workSheet.Column(15).Width = 7.00 + 0.71;
            workSheet.Column(16).Width = 7.14 + 0.71;
            workSheet.Column(17).Width = 7.29  + 0.71;
            workSheet.Column(18).Width = 7.00 + 0.71;
            workSheet.Column(19).Width = 7.00 + 0.71;
            workSheet.Column(20).Width = 5.71 + 0.71;
            workSheet.Column(21).Width = 5.71 + 0.71;


            workSheet.Row(1).Height = 11.25;
            workSheet.Row(2).Height = 20.25;
            workSheet.Row(3).Height = 3.75;
            workSheet.Row(4).Height = 6.75;
            workSheet.Row(5).Height = 10.5;
            workSheet.Row(6).Height = 15.00;
            workSheet.Row(7).Height = 14.25;
            workSheet.Row(8).Height = 15.00;
            workSheet.Row(9).Height = 15.00;
            workSheet.Row(10).Height = 15.00;
            workSheet.Row(11).Height = 6;
            workSheet.Row(12).Height = 22.5;
            workSheet.Row(13).Height = 26.25;
            workSheet.Row(14).Height = 3.00;
            workSheet.Row(15).Height = 24.75;
            workSheet.Row(16).Height = 0;
            workSheet.Row(17).Height = 0;
            workSheet.Row(18).Height = 58.5;
            workSheet.Row(19).Height = 52.5;
            workSheet.Row(19).Height = 41.25;
            
        }
        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A2:U2"].Merge = true;
            workSheet.Cells["A5:H5"].Merge = true;
            workSheet.Cells["K5:U5"].Merge = true;

            workSheet.Cells["A6:D6"].Merge = true;
            workSheet.Cells["E6:H6"].Merge = true;
            workSheet.Cells["L6:U6"].Merge = true;

            workSheet.Cells["A7:D7"].Merge = true;
            workSheet.Cells["E7:H7"].Merge = true;
            workSheet.Cells["L7:U7"].Merge = true;


            workSheet.Cells["A8:D8"].Merge = true;
            workSheet.Cells["E8:H8"].Merge = true;
            workSheet.Cells["K8:M8"].Merge = true;

            workSheet.Cells["N8:U8"].Merge = true;

            workSheet.Cells["A9:D9"].Merge = true;
            workSheet.Cells["E9:H9"].Merge = true;
            workSheet.Cells["K9:M9"].Merge = true;
            workSheet.Cells["N9:U9"].Merge = true;


            workSheet.Cells["A10:D10"].Merge = true;
            workSheet.Cells["G10:H10"].Merge = true;
            workSheet.Cells["K10:M10"].Merge = true;
            workSheet.Cells["N10:U10"].Merge = true;

            workSheet.Cells["A12:U12"].Merge = true;
            workSheet.Cells["A13:U13"].Merge = true;

            //DETALLE

            workSheet.Cells["A14:A20"].Merge = true;
            workSheet.Cells["B14:B20"].Merge = true;
            workSheet.Cells["C14:C20"].Merge = true;
            workSheet.Cells["D14:F15"].Merge = true;

            workSheet.Cells["D16:D20"].Merge = true;
            workSheet.Cells["E16:E20"].Merge = true;
            workSheet.Cells["F16:F20"].Merge = true;
            workSheet.Cells["G14:G20"].Merge = true;
            workSheet.Cells["H14:H20"].Merge = true;
            workSheet.Cells["I14:I20"].Merge = true;
            workSheet.Cells["J14:J20"].Merge = true;
            workSheet.Cells["K14:K20"].Merge = true;
            workSheet.Cells["L14:L20"].Merge = true;

            workSheet.Cells["M14:M20"].Merge = true;
            workSheet.Cells["N14:U17"].Merge = true;

            workSheet.Cells["N18:N20"].Merge = true;
            workSheet.Cells["O18:O20"].Merge = true;
            workSheet.Cells["P18:P20"].Merge = true;
            workSheet.Cells["Q18:Q20"].Merge = true;
            workSheet.Cells["R18:R20"].Merge = true;
            workSheet.Cells["S18:S20"].Merge = true;
            workSheet.Cells["T18:T20"].Merge = true;
            workSheet.Cells["U18:U20"].Merge = true;

        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A5:H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K5:U5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A6:D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E6:H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L6:U6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A7:D7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E7:H7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L7:U7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["A8:D8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E8:H8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K8:M8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L8:U8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A9:D9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E9:H9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K9:M9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L9:U9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["A10:D10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G10:H10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["K10:M10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L10:U10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            
            workSheet.Cells["A12:U12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A13:U13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            //DETALLE

            workSheet.Cells["A14:A20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B14:B20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C14:C20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["D14:F15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D16:D20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E16:E20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F16:F20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G14:G20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H14:H20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I14:I20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J14:J20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K14:K20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L14:L20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M14:M20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["N14:U17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["N18:N20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O18:O20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P18:P20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q18:Q20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R18:R20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S18:S20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["T18:T20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["U18:U20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }
        private static void TextoNegrita(ExcelWorksheet workSheet)
        {

        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A5:H5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E7E6E6"));
            workSheet.Cells["K5:U5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E7E6E6"));

            workSheet.Cells["H14:H20"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C0C0C0"));
            workSheet.Cells["I14:I20"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C0C0C0"));
        }
    }
}
