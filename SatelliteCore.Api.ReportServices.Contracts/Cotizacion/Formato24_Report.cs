using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public class Formato24_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato24_Model cotizacion)
        {
            byte[] file;

            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Almenara");

                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 4);
                imagenUnilene.SetSize(180, 70);

                workSheet.Cells.Style.Font.Name = "Calibri Light";
                workSheet.Cells.Style.Font.Size = 9;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["J4"].Value = "COTIZ Nº ";
                workSheet.Cells["J4"].Style.Font.Size = 9;
                workSheet.Cells["J4"].Style.Font.Name = "Batang";
                workSheet.Cells["J4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J4"].Style.Font.Bold = true;

                workSheet.Cells["L4"].Value = cotizacion.Prov_NroCotizacion + " Unilene S.A.C";
                workSheet.Cells["L4"].Style.Font.Size = 9;
                workSheet.Cells["L4"].Style.Font.Name = "Batang";
                workSheet.Cells["L4"].Style.WrapText = true;
                workSheet.Cells["L4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L4"].Style.Font.Bold = true;
                workSheet.Cells["L4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#333399"));

                workSheet.Cells["A10"].Value = "FORMATO DE COTIZACIÓN Y DECLARACION JURADA DE CUMPLIMIENTO DE ESPECIFICACIONES TECNICAS - COMPRAS <= 08 UITS";
                workSheet.Cells["A10"].Style.Font.Size = 10;
                workSheet.Cells["A10"].Style.Font.Name = "Arial";
                workSheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A10"].Style.Font.Bold = true;
                workSheet.Cells["A10"].Style.WrapText = true;

                workSheet.Cells["A12"].Value = "DATOS DEL PROVEEDOR";
                workSheet.Cells["A12"].Style.Font.Size = 10;
                workSheet.Cells["A12"].Style.Font.Name = "Arial";
                workSheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A12"].Style.Font.Bold = true;
                workSheet.Cells["A12"].Style.WrapText = true;

                workSheet.Cells["J12"].Value = "DATOS DE LA PERSONA DE  CONTACTO DEL POSTOR";
                workSheet.Cells["J12"].Style.Font.Size = 10;
                workSheet.Cells["J12"].Style.Font.Name = "Arial";
                workSheet.Cells["J12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J12"].Style.Font.Bold = true;
                workSheet.Cells["J12"].Style.WrapText = true;

                workSheet.Cells["A13"].Value = "RAZON SOCIAL";
                workSheet.Cells["A13"].Style.Font.Size = 10;
                workSheet.Cells["A13"].Style.Font.Name = "Arial";
                workSheet.Cells["A13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A13"].Style.WrapText = true;

                workSheet.Cells["D13"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["D13"].Style.Font.Size = 10;
                workSheet.Cells["D13"].Style.Font.Name = "Arial";
                workSheet.Cells["D13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D13"].Style.WrapText = true;

                workSheet.Cells["J13"].Value = "NOMBRE";
                workSheet.Cells["J13"].Style.Font.Size = 10;
                workSheet.Cells["J13"].Style.Font.Name = "Arial";
                workSheet.Cells["J13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J13"].Style.WrapText = true;


                workSheet.Cells["K13"].Value = cotizacion.Prov_Contacto;
                workSheet.Cells["K13"].Style.Font.Size = 10;
                workSheet.Cells["K13"].Style.Font.Name = "Arial";
                workSheet.Cells["K13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K13"].Style.WrapText = true;


                workSheet.Cells["A14"].Value = "DIRECCION";
                workSheet.Cells["A14"].Style.Font.Size = 10;
                workSheet.Cells["A14"].Style.Font.Name = "Arial";
                workSheet.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A14"].Style.WrapText = true;

                workSheet.Cells["D14"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["D14"].Style.Font.Size = 10;
                workSheet.Cells["D14"].Style.Font.Name = "Arial";
                workSheet.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D14"].Style.WrapText = true;


                workSheet.Cells["J14"].Value = "CARGO";
                workSheet.Cells["J14"].Style.Font.Size = 10;
                workSheet.Cells["J14"].Style.Font.Name = "Arial";
                workSheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J14"].Style.WrapText = true;


                workSheet.Cells["K14"].Value = cotizacion.Prov_ContCargo;
                workSheet.Cells["K14"].Style.Font.Size = 10;
                workSheet.Cells["K14"].Style.Font.Name = "Arial";
                workSheet.Cells["K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K14"].Style.WrapText = true;


                workSheet.Cells["A15"].Value = "TELEFONO (S) FIJO Y MOVIL";
                workSheet.Cells["A15"].Style.Font.Size = 10;
                workSheet.Cells["A15"].Style.Font.Name = "Arial";
                workSheet.Cells["A15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A15"].Style.WrapText = true;

                workSheet.Cells["D15"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["D15"].Style.Font.Size = 10;
                workSheet.Cells["D15"].Style.Font.Name = "Arial";
                workSheet.Cells["D15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D15"].Style.WrapText = true;


                workSheet.Cells["J15"].Value = "MOVIL (Indicar RPM ó RPC)";
                workSheet.Cells["J15"].Style.Font.Size = 10;
                workSheet.Cells["J15"].Style.Font.Name = "Arial";
                workSheet.Cells["J15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J15"].Style.WrapText = true;


                workSheet.Cells["M15"].Value = cotizacion.Prov_ContCelular;
                workSheet.Cells["M15"].Style.Font.Size = 10;
                workSheet.Cells["M15"].Style.Font.Name = "Arial";
                workSheet.Cells["M15"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M15"].Style.WrapText = true;

                workSheet.Cells["A16"].Value = "EMAIL EMPRESA: ";
                workSheet.Cells["A16"].Style.Font.Size = 10;
                workSheet.Cells["A16"].Style.Font.Name = "Arial";
                workSheet.Cells["A16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A16"].Style.WrapText = true;

                workSheet.Cells["D16"].Value = cotizacion.Prov_Email;
                workSheet.Cells["D16"].Style.Font.Size = 10;
                workSheet.Cells["D16"].Style.Font.Name = "Arial";
                workSheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D16"].Style.WrapText = true;

                workSheet.Cells["J16"].Value = "EMAIL DEL CONTACTO ";
                workSheet.Cells["J16"].Style.Font.Size = 10;
                workSheet.Cells["J16"].Style.Font.Name = "Arial";
                workSheet.Cells["J16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J16"].Style.WrapText = true;

                workSheet.Cells["M16"].Value = cotizacion.Prov_ContEmail;
                workSheet.Cells["M16"].Style.Font.Size = 10;
                workSheet.Cells["M16"].Style.Font.Name = "Arial";
                workSheet.Cells["M16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M16"].Style.WrapText = true;


                workSheet.Cells["A18"].Value = "VIGENCIA DE LA OFERTA";
                workSheet.Cells["A18"].Style.Font.Size = 10;
                workSheet.Cells["A18"].Style.Font.Name = "Arial";
                workSheet.Cells["A18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A18"].Style.WrapText = true;

                workSheet.Cells["D18"].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["D18"].Style.Font.Size = 10;
                workSheet.Cells["D18"].Style.Font.Name = "Arial";
                workSheet.Cells["D18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D18"].Style.WrapText = true;



                workSheet.Cells["A19"].Value = "PLAZO DE ENTREGA  (De acuero a lo indicado en los requerimientos técnicos mínimos)";
                workSheet.Cells["A19"].Style.Font.Size = 10;
                workSheet.Cells["A19"].Style.Font.Name = "Arial";
                workSheet.Cells["A19"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A19"].Style.WrapText = true;

                workSheet.Cells["D19"].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["D19"].Style.Font.Size = 10;
                workSheet.Cells["D19"].Style.Font.Name = "Arial";
                workSheet.Cells["D19"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D19"].Style.WrapText = true;



                workSheet.Cells["A20"].Value = "GARANTIA (De acuero a lo indicado en los requerimientos técnicos mínimos)";
                workSheet.Cells["A20"].Style.Font.Size = 10;
                workSheet.Cells["A20"].Style.Font.Name = "Arial";
                workSheet.Cells["A20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A20"].Style.WrapText = true;

                workSheet.Cells["D20"].Value = cotizacion.Prov_Garantia;
                workSheet.Cells["D20"].Style.Font.Size = 10;
                workSheet.Cells["D20"].Style.Font.Name = "Arial";
                workSheet.Cells["D20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D20"].Style.WrapText = true;


                workSheet.Cells["A21"].Value = "FECHA  DE  LA OFERTA";
                workSheet.Cells["A21"].Style.Font.Size = 10;
                workSheet.Cells["A21"].Style.Font.Name = "Arial";
                workSheet.Cells["A21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A21"].Style.WrapText = true;


                workSheet.Cells["C21"].Value = cotizacion.Prov_FechaOferta;
                workSheet.Cells["C21"].Style.Font.Size = 10;
                workSheet.Cells["C21"].Style.Font.Name = "Arial";
                workSheet.Cells["C21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C21"].Style.WrapText = true;
                workSheet.Cells["C21"].Style.Numberformat.Format = "dd/MM/yyyy";

                workSheet.Cells["E21"].Value ="RUC";
                workSheet.Cells["E21"].Style.Font.Size = 10;
                workSheet.Cells["E21"].Style.Font.Name = "Arial";
                workSheet.Cells["E21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E21"].Style.WrapText = true;

                workSheet.Cells["F21"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["F21"].Style.Font.Size = 10;
                workSheet.Cells["F21"].Style.Font.Name = "Arial";
                workSheet.Cells["F21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F21"].Style.WrapText = true;


                workSheet.Cells["A24"].Value = "Señores:  ESSALUD";
                workSheet.Cells["A24"].Style.Font.Bold = true;
                workSheet.Cells["A24"].Style.Font.Size = 10;
                workSheet.Cells["A24"].Style.Font.Name = "Arial";
                workSheet.Cells["A24"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A24"].Style.WrapText = true;

                workSheet.Cells["A25"].Value = "De mi consideración:";
                workSheet.Cells["A25"].Style.Font.Size = 10;
                workSheet.Cells["A25"].Style.Font.Name = "Arial";
                workSheet.Cells["A25"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A25"].Style.WrapText = true;


                workSheet.Cells["A27"].Value = "En respuesta a la solicitud de cotización sobre la 'Adquisición de Material Médico  Delegado a Compra Local' y despues de haber analizado las especificaciones técnicas del mencionado requerimiento, " +
                    "las mismas que acepto en todos sus extremos, indico que cumplo con todos los requerimientos solicitados.";
                workSheet.Cells["A27"].Style.Font.Size = 10;
                workSheet.Cells["A27"].Style.Font.Bold = true;
                workSheet.Cells["A27"].Style.Font.Name = "Arial";
                workSheet.Cells["A27"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A27"].Style.WrapText = true;

                workSheet.Cells["A28"].Value = "Asimismo, declaro que las características técnicas de los bienes cotizados por mi representada se ajustan" +
                    " a lo requerido por su totalidad. En tal sentido, indico que el costo";
                workSheet.Cells["A28"].Style.Font.Size = 10;
                workSheet.Cells["A28"].Style.Font.Name = "Arial";
                workSheet.Cells["A28"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A28"].Style.Font.Bold = true;
                workSheet.Cells["A28"].Style.WrapText = true;





                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);


                //DETALLE 


                workSheet.Cells["A29"].Value = "Item";
                workSheet.Cells["A29"].Style.Font.Size = 9;
                workSheet.Cells["A29"].Style.Font.Name = "Arial Narrow";
                workSheet.Cells["A29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A29"].Style.WrapText = true;


                workSheet.Cells["B29"].Value = "SOLPED";
                workSheet.Cells["B29"].Style.Font.Size = 9;
                workSheet.Cells["B29"].Style.Font.Name = "Arial";
                workSheet.Cells["B29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B29"].Style.WrapText = true;

                workSheet.Cells["C29"].Value = "DESCRIPCIÓN DEL ITEM";
                workSheet.Cells["C29"].Style.Font.Size = 9;
                workSheet.Cells["C29"].Style.Font.Name = "Arial";
                workSheet.Cells["C29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C29"].Style.Font.Bold = true;
                workSheet.Cells["C29"].Style.WrapText = true;

                workSheet.Cells["C32"].Value = "CODIGO SAP";
                workSheet.Cells["C32"].Style.Font.Size = 9;
                workSheet.Cells["C32"].Style.Font.Name = "Arial";
                workSheet.Cells["C32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C32"].Style.WrapText = true;

                workSheet.Cells["D32"].Value = "PRODUCTO (ITEM)";
                workSheet.Cells["D32"].Style.Font.Size = 9;
                workSheet.Cells["D32"].Style.Font.Name = "Arial";
                workSheet.Cells["D32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D32"].Style.WrapText = true;

                workSheet.Cells["E32"].Value = "UM";
                workSheet.Cells["E32"].Style.Font.Size = 9;
                workSheet.Cells["E32"].Style.Font.Name = "Arial";
                workSheet.Cells["E32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E32"].Style.WrapText = true;

                workSheet.Cells["F29"].Value = "Cantidad Total";
                workSheet.Cells["F29"].Style.Font.Size = 9;
                workSheet.Cells["F29"].Style.Font.Name = "Arial";
                workSheet.Cells["F29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F29"].Style.WrapText = true;

                workSheet.Cells["G29"].Value = "Precio Unitario S/.";
                workSheet.Cells["G29"].Style.Font.Size = 9;
                workSheet.Cells["G29"].Style.Font.Name = "Arial";
                workSheet.Cells["G29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G29"].Style.WrapText = true;


                workSheet.Cells["H29"].Value = "Precio Total S/.";
                workSheet.Cells["H29"].Style.Font.Size = 9;
                workSheet.Cells["H29"].Style.Font.Name = "Arial";
                workSheet.Cells["H29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H29"].Style.WrapText = true;


                workSheet.Cells["I29"].Value = "Marca";
                workSheet.Cells["I29"].Style.Font.Size = 9;
                workSheet.Cells["I29"].Style.Font.Name = "Arial";
                workSheet.Cells["I29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I29"].Style.WrapText = true;

                workSheet.Cells["J29"].Value = "Procedencia";
                workSheet.Cells["J29"].Style.Font.Size = 9;
                workSheet.Cells["J29"].Style.Font.Name = "Arial";
                workSheet.Cells["J29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J29"].Style.WrapText = true;

                

                workSheet.Cells["K29"].Value = "Forma de Presentación del Producto";
                workSheet.Cells["K29"].Style.Font.Size = 9;
                workSheet.Cells["K29"].Style.Font.Name = "Arial";
                workSheet.Cells["K29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K29"].Style.WrapText = true;

                workSheet.Cells["L29"].Value = "Plazo de Entrega total del Producto";
                workSheet.Cells["L29"].Style.Font.Size = 9;
                workSheet.Cells["L29"].Style.Font.Name = "Arial";
                workSheet.Cells["L29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L29"].Style.WrapText = true;

                workSheet.Cells["M29"].Value = "DECLARACION JURADA DE CUMPLIMIENTO DE EE.TT. MINIMAS";
                workSheet.Cells["M29"].Style.Font.Size = 9;
                workSheet.Cells["M29"].Style.Font.Name = "Arial";
                workSheet.Cells["M29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M29"].Style.WrapText = true;

                workSheet.Cells["M32"].Value = "Vigencia del material médico no menor de 18 meses";
                workSheet.Cells["M32"].Style.Font.Size = 9;
                workSheet.Cells["M32"].Style.Font.Name = "Arial";
                workSheet.Cells["M32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M32"].Style.WrapText = true;
                workSheet.Cells["M32"].Style.TextRotation = 90;



                workSheet.Cells["N32"].Value = "Capacidad de Atención al 100% de la Cantdidad Total solicitada (Si o No)";
                workSheet.Cells["N32"].Style.Font.Size = 9;
                workSheet.Cells["N32"].Style.Font.Name = "Arial";
                workSheet.Cells["N32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N32"].Style.WrapText = true;
                workSheet.Cells["N32"].Style.TextRotation = 90;

                workSheet.Cells["O32"].Value = "Cumple con los Plazos de Entrega establecidos en su cotización (Si  o No)";
                workSheet.Cells["O32"].Style.Font.Size = 9;
                workSheet.Cells["O32"].Style.Font.Name = "Arial";
                workSheet.Cells["O32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O32"].Style.WrapText = true;
                workSheet.Cells["O32"].Style.TextRotation = 90;


                workSheet.Cells["P32"].Value = "Certificado de Buenas Prácticas de Almacenamiento igente(Si o No)";
                workSheet.Cells["P32"].Style.Font.Size = 9;
                workSheet.Cells["P32"].Style.Font.Name = "Arial";
                workSheet.Cells["P32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P32"].Style.WrapText = true;
                workSheet.Cells["P32"].Style.TextRotation = 90;

                workSheet.Cells["Q32"].Value = "Certificado de Buenas Prácticas de Almacenamiento igente(Si o No)";
                workSheet.Cells["Q32"].Style.Font.Size = 9;
                workSheet.Cells["Q32"].Style.Font.Name = "Arial";
                workSheet.Cells["Q32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q32"].Style.WrapText = true;
                workSheet.Cells["Q32"].Style.TextRotation = 90;

                workSheet.Cells["R32"].Value = "Certificado de Buenas Prácticas de Almacenamiento igente(Si o No)";
                workSheet.Cells["R32"].Style.Font.Size = 9;
                workSheet.Cells["R32"].Style.Font.Name = "Arial";
                workSheet.Cells["R32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R32"].Style.WrapText = true;
                workSheet.Cells["R32"].Style.TextRotation = 90;

                workSheet.Cells["S32"].Value = "Cumple con las normas de conservación y transporte del producto y se compromete a canjear de haber alguna observación a la recepción del producto  (Sí o No)";
                workSheet.Cells["S32"].Style.Font.Size = 9;
                workSheet.Cells["S32"].Style.Font.Name = "Arial";
                workSheet.Cells["S32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S32"].Style.WrapText = true;
                workSheet.Cells["S32"].Style.TextRotation = 90;

                int index = 35;

                foreach (Coti_Formato24_Detalle item in cotizacion.Detalle)
                {

                    workSheet.Row(index).Height = 59.25;
                    
                    workSheet.Cells["A" + index].Value = index - 34;
                    workSheet.Cells["A" + index].Style.Font.Size = 8;
                    workSheet.Cells["A" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";


                    workSheet.Cells["B" + index].Value = item.Solped;
                    workSheet.Cells["B" + index].Style.Font.Size = 8;
                    workSheet.Cells["B" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                   

                    workSheet.Cells["C" + index].Value = item.CodigoSap;
                    workSheet.Cells["C" + index].Style.Font.Size = 8;
                    workSheet.Cells["C" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = item.Producto;
                    workSheet.Cells["D" + index].Style.Font.Size = 8;
                    workSheet.Cells["D" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = item.Um;
                    workSheet.Cells["E" + index].Style.Font.Size = 8;
                    workSheet.Cells["E" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = item.Cantidad;
                    workSheet.Cells["F" + index].Style.Font.Size = 8;
                    workSheet.Cells["F" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["G" + index].Value = item.PreUnitario;
                    workSheet.Cells["G" + index].Style.Font.Size = 8;
                    workSheet.Cells["G" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["H" + index].Value = item.PreTotal;
                    workSheet.Cells["H" + index].Style.Font.Size = 8;
                    workSheet.Cells["H" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["I" + index].Value = item.Marca;
                    workSheet.Cells["I" + index].Style.Font.Size = 8;
                    workSheet.Cells["I" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;


                    workSheet.Cells["J" + index].Value = item.Procedencia;
                    workSheet.Cells["J" + index].Style.Font.Size = 8;
                    workSheet.Cells["J" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;


                    workSheet.Cells["K" + index].Value = item.Presentacion;
                    workSheet.Cells["K" + index].Style.Font.Size = 8;
                    workSheet.Cells["K" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;


                    workSheet.Cells["L" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["L" + index].Style.Font.Size = 8;
                    workSheet.Cells["L" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;

                    workSheet.Cells["M" + index].Value = item.VigMaterial;
                    workSheet.Cells["M" + index].Style.Font.Size = 8;
                    workSheet.Cells["M" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    workSheet.Cells["N" + index].Value = item.CapAtencion;
                    workSheet.Cells["N" + index].Style.Font.Size = 8;
                    workSheet.Cells["N" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;


                    workSheet.Cells["O" + index].Value = item.CumpPlazo;
                    workSheet.Cells["O" + index].Style.Font.Size = 8;
                    workSheet.Cells["O" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;

                    workSheet.Cells["P" + index].Value = item.CBuenasPracticas;
                    workSheet.Cells["P" + index].Style.Font.Size = 8;
                    workSheet.Cells["P" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["P" + index].Style.WrapText = true;

                    workSheet.Cells["Q" + index].Value = item.RegSanitario;
                    workSheet.Cells["Q" + index].Style.Font.Size = 8;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;


                    workSheet.Cells["R" + index].Value = item.CertBPracticas;
                    workSheet.Cells["R" + index].Style.Font.Size = 8;
                    workSheet.Cells["R" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    workSheet.Cells["S" + index].Value = item.CumpNormas;
                    workSheet.Cells["S" + index].Style.Font.Size = 8;
                    workSheet.Cells["S" + index].Style.Font.Name = "Calibri";
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    index++;

                }


                workSheet.Cells["A" + index + ":G" + index].Merge = true;
                workSheet.Cells["A" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index].Value = "VALOR TOTAL DE LA COTIZACON S/";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["H" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["H" + index].Style.Numberformat.Format = "#,##0.00";


                index++;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "CONSIDERACIONES GENERALES: ";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;

                index++;
                workSheet.Row(index).Height = 47.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye totos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso, los costos laborales conforme la legislación vigente, así como cualquier otro concepto que peda tener incidenca sobre el costo del bien y/o servicio a contratar, excepto la de aquellos proveeores que gocen de alguna exoneración legal, no incluiran en el precio de su oferta los tributos respectivos. Asimismo, declaro bajo juramento que, mi persona y/o representada no cuenta con impedimentos para contratar con el Estado, conforme lo establece el artículo 11 del Texto Único Ordenado de la Ley Nº 30225, Ley de Contrataciones del Estado, aprobado por Declreto N. 082-2019-EF.";
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;


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
            workSheet.Column(1).Width = 7.86 + 0.71;
            workSheet.Column(2).Width = 7.86 + 0.71;
            workSheet.Column(3).Width = 12.57 + 0.71;
            workSheet.Column(4).Width = 24.86 + 0.71;
            workSheet.Column(5).Width = 4.57 + 0.71; 
            workSheet.Column(6).Width = 7.29 + 0.71;
            workSheet.Column(7).Width = 8.57 + 0.71;
            workSheet.Column(8).Width = 10.43 + 0.71;
            workSheet.Column(9).Width = 10.43 + 0.71;
            workSheet.Column(10).Width = 8.57 + 0.71;
            workSheet.Column(11).Width = 7.86 + 0.71;
            workSheet.Column(12).Width = 9.57 + 0.71;
            workSheet.Column(13).Width = 7.29 + 0.71;
            workSheet.Column(14).Width = 7.29 + 0.71;
            workSheet.Column(15).Width = 7 + 0.71;
            workSheet.Column(16).Width = 7 + 0.71;
            workSheet.Column(17).Width = 5.71 + 0.71;
            workSheet.Column(18).Width = 5.14 + 0.71;
            workSheet.Column(19).Width = 12.86 + 0.71;

            workSheet.Row(1).Height = 15;
            workSheet.Row(2).Height = 11.25;
            workSheet.Row(3).Height = 3.75;
            workSheet.Row(4).Height = 15.75;
            workSheet.Row(5).Height = 3.75;
            workSheet.Row(6).Height = 3.75;
            workSheet.Row(7).Height = 3.75;
            workSheet.Row(8).Height = 3.75;
            workSheet.Row(9).Height = 11.25;
            workSheet.Row(10).Height = 21;
            workSheet.Row(11).Height = 4.5;
            workSheet.Row(12).Height = 10.5;
            workSheet.Row(13).Height = 11.25;
            workSheet.Row(14).Height = 12;
            workSheet.Row(15).Height = 12;
            workSheet.Row(16).Height = 15.75;
            workSheet.Row(17).Height = 7.5;
            workSheet.Row(18).Height = 13.5;
            workSheet.Row(19).Height = 25.5;
            workSheet.Row(20).Height = 36;
            workSheet.Row(21).Height = 18.75;
            workSheet.Row(22).Height = 6.00;
            workSheet.Row(23).Height = 3.00;

            workSheet.Row(24).Height = 20.25;
            workSheet.Row(25).Height = 19.5;
            workSheet.Row(26).Height = 6;
            workSheet.Row(27).Height = 31.5;
            workSheet.Row(28).Height = 31.5;


            workSheet.Row(29).Height = 15.75;

            workSheet.Row(30).Height = 00.00;
            workSheet.Row(31).Height = 00.00;
            workSheet.Row(32).Height = 59.25;
            workSheet.Row(33).Height = 22.50;
            workSheet.Row(34).Height = 35.25;

            
                






        }

        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["J4:K4"].Merge = true;
            workSheet.Cells["L4:O4"].Merge = true;
            workSheet.Cells["A10:S10"].Merge = true;

            workSheet.Cells["A12:G12"].Merge = true;
            workSheet.Cells["J12:S12"].Merge = true;

            workSheet.Cells["A13:C13"].Merge = true;
            workSheet.Cells["D13:G13"].Merge = true;
            workSheet.Cells["K13:S13"].Merge = true;

            workSheet.Cells["A14:C14"].Merge = true;
            workSheet.Cells["D14:G14"].Merge = true;
            workSheet.Cells["K14:S14"].Merge = true;

            workSheet.Cells["A15:C15"].Merge = true;
            workSheet.Cells["D15:G15"].Merge = true;
            workSheet.Cells["J15:L15"].Merge = true;
            workSheet.Cells["M15:S15"].Merge = true;

            workSheet.Cells["A16:C17"].Merge = true;
            workSheet.Cells["D16:G17"].Merge = true;

            workSheet.Cells["J16:L16"].Merge = true;
            workSheet.Cells["M16:S16"].Merge = true;


            workSheet.Cells["A18:C18"].Merge = true;
            workSheet.Cells["D18:G18"].Merge = true;
            workSheet.Cells["A19:C19"].Merge = true;
            workSheet.Cells["D19:G19"].Merge = true;
            workSheet.Cells["A20:C20"].Merge = true;
            workSheet.Cells["D20:G20"].Merge = true;

            workSheet.Cells["A21:C21"].Merge = true;
            workSheet.Cells["F21:G21"].Merge = true;

            workSheet.Cells["A24:C24"].Merge = true;

            workSheet.Cells["A25:C25"].Merge = true;

            workSheet.Cells["A27:S27"].Merge = true;
            workSheet.Cells["A28:S28"].Merge = true;

            workSheet.Cells["A29:A34"].Merge = true;
            workSheet.Cells["B29:B34"].Merge = true;
            workSheet.Cells["C29:E31"].Merge = true;

            workSheet.Cells["C32:C34"].Merge = true;
            workSheet.Cells["D32:D34"].Merge = true;
            workSheet.Cells["E32:E34"].Merge = true;
            workSheet.Cells["F29:F34"].Merge = true;
            workSheet.Cells["G29:G34"].Merge = true;

            workSheet.Cells["H29:H34"].Merge = true;
            workSheet.Cells["I29:I34"].Merge = true;
            workSheet.Cells["J29:J34"].Merge = true;
            workSheet.Cells["K29:K34"].Merge = true;
            workSheet.Cells["L29:L34"].Merge = true;
            workSheet.Cells["M29:S31"].Merge = true;


            workSheet.Cells["M32:M34"].Merge = true;
            workSheet.Cells["N32:N34"].Merge = true;
            workSheet.Cells["O32:O34"].Merge = true;
            workSheet.Cells["P32:P34"].Merge = true;
            workSheet.Cells["Q32:Q34"].Merge = true;
            workSheet.Cells["R32:R34"].Merge = true;
            workSheet.Cells["S32:S34"].Merge = true;

        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["J4:K4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L4:O4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A10:S10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A12:G12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J12:S12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A13:C13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D13:G13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K13:S13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A14:C14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D14:G14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K14:S14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["A15:C15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D15:G15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["J15:L15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M15:S15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A16:C17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D16:G17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["J16:L16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M16:S16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A18:C18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D18:G18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["A19:C19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D19:G19"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A20:C20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D20:G20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A21:C21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F21:G21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A23:S28"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["A29:A34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B29:B34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C29:E31"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C32:C34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D32:D34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E32:E34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F29:F34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G29:G34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H29:H34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I29:I34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J29:J34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K29:K34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L29:L34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            workSheet.Cells["M29:S31"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            workSheet.Cells["M32:M34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N32:N34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O32:O34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P32:P34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q32:Q34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R32:R34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S32:S34"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }

        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A29:S34"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B7DEE8"));
            //workSheet.Cells["B29:B34"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B7DEE8"));
           // workSheet.Cells["C29:D31"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B7DEE8"));
        }

        private static void TextoNegrita(ExcelWorksheet workSheet) { }

    }
}
