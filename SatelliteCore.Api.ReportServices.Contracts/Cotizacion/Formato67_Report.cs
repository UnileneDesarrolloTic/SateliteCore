using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato67_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Coti_Formato_67_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("EsSalud Nacional del Niño San Borja");
                ExcelPicture imagenUnilene = worksheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(220, 60);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                PintarCeldas(worksheet);
                BordesCeldas(worksheet);
                BordesCeldas(worksheet);
                TextoNegrita(worksheet);




                worksheet.Cells["L4"].Value = "COTIZ NRO. " + cotizacion.Prov_NroCotizacion + "-2022 Unilene S.A.C";
                worksheet.Cells["L4"].Style.Font.Size = 12;
                worksheet.Cells["L4"].Style.Font.Name = "Calibri";
                worksheet.Cells["L4"].Style.Font.Bold = true;
                worksheet.Cells["L4"].Style.WrapText = true;
                worksheet.Cells["L4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                //worksheet.Cells["L4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                worksheet.Cells["A5"].Value = "SOLICITUD DE COTIZACIÓN";
                worksheet.Cells["A5"].Style.Font.Size = 16;
                worksheet.Cells["A5"].Style.Font.Name = "Arial";
                worksheet.Cells["A5"].Style.Font.Bold = true;
                worksheet.Cells["A5"].Style.WrapText = true;
                worksheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N5"].Value = "Fecha de Emisión";
                worksheet.Cells["N5"].Style.Font.Size = 10;
                worksheet.Cells["N5"].Style.Font.Name = "Arial";
                worksheet.Cells["N5"].Style.WrapText = true;
                worksheet.Cells["N5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A6"].Value = "SUB UNIDAD DE LOGISTICA";
                worksheet.Cells["A6"].Style.Font.Size = 11;
                worksheet.Cells["A6"].Style.Font.Name = "Arial";
                worksheet.Cells["A6"].Style.Font.Bold = true;
                worksheet.Cells["A6"].Style.WrapText = true;
                worksheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["N6"].Value = cotizacion.Prov_Fecha;
                worksheet.Cells["N6"].Style.Font.Size = 10;
                worksheet.Cells["N6"].Style.Font.Name = "Arial";
                worksheet.Cells["N6"].Style.WrapText = true;
                worksheet.Cells["N6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N6"].Style.Numberformat.Format = "dd/MM/yyyy";


                worksheet.Cells["N7"].Value = "Validez de la Oferta";
                worksheet.Cells["N7"].Style.Font.Size = 10;
                worksheet.Cells["N7"].Style.Font.Name = "Arial";
                worksheet.Cells["N7"].Style.WrapText = true;
                worksheet.Cells["N7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                worksheet.Cells["N8"].Value = cotizacion.Prov_vigOferta;
                worksheet.Cells["N8"].Style.Font.Size = 10;
                worksheet.Cells["N8"].Style.Font.Name = "Arial";
                worksheet.Cells["N8"].Style.WrapText = true;
                worksheet.Cells["N8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A10"].Value = "SR.PROVEEDOR:";
                worksheet.Cells["A10"].Style.Font.Size = 10;
                worksheet.Cells["A10"].Style.Font.Name = "Arial";
                worksheet.Cells["A10"].Style.WrapText = true;
                worksheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A11"].Value = "Por medio de la presente solicitamos se sirva llenar la solcitud de cotización para precios referenciales de lo siguiente:(en concordancia al " +
                    "Articulo 13º del Reglamento de la Ley de Contrataciones del Estado; es decir que su valor referencial deberá ser calculado a todo costo, incluyendo todos los tributos, seguros, transportes, inspecciones, pruebas y, " +
                    "de ser el caso , los costos laborales respectivos conforme a la legislación vigente,a sí como cualquier otro conscepto que puieda incidir sobre el costo de los bienes).";
                worksheet.Cells["A11"].Style.Font.Size = 11;
                worksheet.Cells["A11"].Style.Font.Name = "Arial";
                worksheet.Cells["A11"].Style.WrapText = true;
                worksheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                worksheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                worksheet.Cells["A12"].Value = "RAZÓN SOCIAL:";
                worksheet.Cells["A12"].Style.Font.Size = 10;
                worksheet.Cells["A12"].Style.Font.Name = "Arial";
                worksheet.Cells["A12"].Style.WrapText = true;
                worksheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A12"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["B12"].Value = cotizacion.Prov_RazonSocial;
                worksheet.Cells["B12"].Style.Font.Size = 10;
                worksheet.Cells["B12"].Style.Font.Name = "Arial";
                worksheet.Cells["B12"].Style.WrapText = true;
                worksheet.Cells["B12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
              

                worksheet.Cells["M12"].Value = "RUC:";
                worksheet.Cells["M12"].Style.Font.Size = 10;
                worksheet.Cells["M12"].Style.Font.Name = "Arial";
                worksheet.Cells["M12"].Style.WrapText = true;
                worksheet.Cells["M12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["M12"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["N12"].Value =cotizacion.Prov_Ruc;
                worksheet.Cells["N12"].Style.Font.Size = 10;
                worksheet.Cells["N12"].Style.Font.Name = "Arial";
                worksheet.Cells["N12"].Style.WrapText = true;
                worksheet.Cells["N12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                worksheet.Cells["A13"].Value = "DIRECCIÓN:";
                worksheet.Cells["A13"].Style.Font.Size = 10;
                worksheet.Cells["A13"].Style.Font.Name = "Arial";
                worksheet.Cells["A13"].Style.WrapText = true;
                worksheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A13"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["B13"].Value = cotizacion.Prov_Direccion;
                worksheet.Cells["B13"].Style.Font.Size = 10;
                worksheet.Cells["B13"].Style.Font.Name = "Arial";
                worksheet.Cells["B13"].Style.WrapText = true;
                worksheet.Cells["B13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A14"].Value = "TELEFONO:";
                worksheet.Cells["A14"].Style.Font.Size = 10;
                worksheet.Cells["A14"].Style.Font.Name = "Arial";
                worksheet.Cells["A14"].Style.WrapText = true;
                worksheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A14"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["B14"].Value = cotizacion.Prov_Telefono;
                worksheet.Cells["B14"].Style.Font.Size = 10;
                worksheet.Cells["B14"].Style.Font.Name = "Arial";
                worksheet.Cells["B14"].Style.WrapText = true;
                worksheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A15"].Value = "CONTACTO:";
                worksheet.Cells["A15"].Style.Font.Size = 10;
                worksheet.Cells["A15"].Style.Font.Name = "Arial";
                worksheet.Cells["A15"].Style.WrapText = true;
                worksheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A15"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["B15"].Value = cotizacion.Prov_Contacto;
                worksheet.Cells["B15"].Style.Font.Size = 10;
                worksheet.Cells["B15"].Style.Font.Name = "Arial";
                worksheet.Cells["B15"].Style.WrapText = true;
                worksheet.Cells["B15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A16"].Value = "E-MAIL:";
                worksheet.Cells["A16"].Style.Font.Size = 10;
                worksheet.Cells["A16"].Style.Font.Name = "Arial";
                worksheet.Cells["A16"].Style.WrapText = true;
                worksheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A16"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["B16"].Value = cotizacion.Prov_Email;
                worksheet.Cells["B16"].Style.Font.Size = 10;
                worksheet.Cells["B16"].Style.Font.Name = "Arial";
                worksheet.Cells["B16"].Style.WrapText = true;
                worksheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A17"].Value = "CELULAR:";
                worksheet.Cells["A17"].Style.Font.Size = 10;
                worksheet.Cells["A17"].Style.Font.Name = "Arial";
                worksheet.Cells["A17"].Style.WrapText = true;
                worksheet.Cells["A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A17"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["B17"].Value = cotizacion.Prov_movil;
                worksheet.Cells["B17"].Style.Font.Size = 10;
                worksheet.Cells["B17"].Style.Font.Name = "Arial";
                worksheet.Cells["B17"].Style.WrapText = true;
                worksheet.Cells["B17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                worksheet.Cells["G17"].Value = "SÍ / NO";
                worksheet.Cells["G17"].Style.Font.Size = 10;
                worksheet.Cells["G17"].Style.Font.Name = "Arial";
                worksheet.Cells["G17"].Style.WrapText = true;
                worksheet.Cells["G17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["H17"].Value = "SÍ / NO";
                worksheet.Cells["H17"].Style.Font.Size = 10;
                worksheet.Cells["H17"].Style.Font.Name = "Arial";
                worksheet.Cells["H17"].Style.WrapText = true;
                worksheet.Cells["H17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J17"].Value = "SÍ / NO";
                worksheet.Cells["J17"].Style.Font.Size = 10;
                worksheet.Cells["J17"].Style.Font.Name = "Arial";
                worksheet.Cells["J17"].Style.WrapText = true;
                worksheet.Cells["J17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["K17"].Value = "SÍ / NO";
                worksheet.Cells["K17"].Style.Font.Size = 10;
                worksheet.Cells["K17"].Style.Font.Name = "Arial";
                worksheet.Cells["K17"].Style.WrapText = true;
                worksheet.Cells["K17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["N17"].Value = "SÍ / NO";
                worksheet.Cells["N17"].Style.Font.Size = 10;
                worksheet.Cells["N17"].Style.Font.Name = "Arial";
                worksheet.Cells["N17"].Style.WrapText = true;
                worksheet.Cells["N17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                //DETALLE


                worksheet.Cells["A18"].Value = "ITEM";
                worksheet.Cells["A18"].Style.Font.Size = 8;
                worksheet.Cells["A18"].Style.Font.Name = "Arial";
                worksheet.Cells["A18"].Style.WrapText = true;
                worksheet.Cells["A18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                worksheet.Cells["B18"].Value = "DESCRIPCIÓN";
                worksheet.Cells["B18"].Style.Font.Size = 8;
                worksheet.Cells["B18"].Style.Font.Name = "Arial";
                worksheet.Cells["B18"].Style.WrapText = true;
                worksheet.Cells["B18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["C18"].Value = "U.M.";
                worksheet.Cells["C18"].Style.Font.Size = 8;
                worksheet.Cells["C18"].Style.Font.Name = "Arial";
                worksheet.Cells["C18"].Style.WrapText = true;
                worksheet.Cells["C18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                worksheet.Cells["D18"].Value = "CANTIDAD";
                worksheet.Cells["D18"].Style.Font.Size = 8;
                worksheet.Cells["D18"].Style.Font.Name = "Arial";
                worksheet.Cells["D18"].Style.WrapText = true;
                worksheet.Cells["D18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                worksheet.Cells["E18"].Value = "PRECIO UNITARIO S/. (INC.IGV.)";
                worksheet.Cells["E18"].Style.Font.Size = 8;
                worksheet.Cells["E18"].Style.Font.Name = "Arial";
                worksheet.Cells["E18"].Style.WrapText = true;
                worksheet.Cells["E18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["F18"].Value = "PRECIO TOTAL S/. (INCL.IGV)";
                worksheet.Cells["F18"].Style.Font.Size = 8;
                worksheet.Cells["F18"].Style.Font.Name = "Arial";
                worksheet.Cells["F18"].Style.WrapText = true;
                worksheet.Cells["F18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                worksheet.Cells["G18"].Value = "RNP VIGENTE";
                worksheet.Cells["G18"].Style.Font.Size = 8;
                worksheet.Cells["G18"].Style.Font.Name = "Arial";
                worksheet.Cells["G18"].Style.WrapText = true;
                worksheet.Cells["G18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));



                worksheet.Cells["H18"].Value = "Nº DE REGISTRO SANITARIO VIGENTE Y LEGIBLE";
                worksheet.Cells["H18"].Style.Font.Size = 8;
                worksheet.Cells["H18"].Style.Font.Name = "Arial";
                worksheet.Cells["H18"].Style.WrapText = true;
                worksheet.Cells["H18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                worksheet.Cells["I18"].Value = "VIGENCIA DEL PRODUCTO COTIZADO DE ACUERDO A LAS E.E.T.T.";
                worksheet.Cells["I18"].Style.Font.Size = 8;
                worksheet.Cells["I18"].Style.Font.Name = "Arial";
                worksheet.Cells["I18"].Style.WrapText = true;
                worksheet.Cells["I18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["J18"].Value = "CERTIFICADO DE BPM Y BPA";
                worksheet.Cells["J18"].Style.Font.Size = 8;
                worksheet.Cells["J18"].Style.Font.Name = "Arial";
                worksheet.Cells["J18"].Style.WrapText = true;
                worksheet.Cells["J18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["K18"].Value = "PROTOCOLO DE ANÁLISIS";
                worksheet.Cells["K18"].Style.Font.Size = 8;
                worksheet.Cells["K18"].Style.Font.Name = "Arial";
                worksheet.Cells["K18"].Style.WrapText = true;
                worksheet.Cells["K18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["L18"].Value = "VALIDEZ DE OFERTA DÍAS";
                worksheet.Cells["L18"].Style.Font.Size = 8;
                worksheet.Cells["L18"].Style.Font.Name = "Arial";
                worksheet.Cells["L18"].Style.WrapText = true;
                worksheet.Cells["L18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["M18"].Value = "MARCA/PAIS PROCEDENCIA";
                worksheet.Cells["M18"].Style.Font.Size = 8;
                worksheet.Cells["M18"].Style.Font.Name = "Arial";
                worksheet.Cells["M18"].Style.WrapText = true;
                worksheet.Cells["M18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                worksheet.Cells["N18"].Value = "CUMPLE CON EL 100% DE LAS EE.TT.";
                worksheet.Cells["N18"].Style.Font.Size = 8;
                worksheet.Cells["N18"].Style.Font.Name = "Arial";
                worksheet.Cells["N18"].Style.WrapText = true;
                worksheet.Cells["N18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                worksheet.Cells["O18"].Value = "PLAZO DE ENTREGA EN DÍAS";
                worksheet.Cells["O18"].Style.Font.Size = 8;
                worksheet.Cells["O18"].Style.Font.Name = "Arial";
                worksheet.Cells["O18"].Style.WrapText = true;
                worksheet.Cells["O18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O18"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));


                int row = 19;

                foreach (Coti_Formato67_Detalle item in cotizacion.Detalle)
                {
                    worksheet.Row(row).Height = 92.25;
                    worksheet.Cells["A"+ row ].Value = item.Item;
                    worksheet.Cells["A" + row].Style.Font.Size = 8;
                    worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["A" + row].Style.WrapText = true;

                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["B" + row].Value = item.Descripcion;
                    worksheet.Cells["B" + row].Style.Font.Size = 12;
                    worksheet.Cells["B" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["B" + row].Style.WrapText = true;

                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    worksheet.Cells["C" + row].Value = item.Um;
                    worksheet.Cells["C" + row].Style.Font.Size = 8;
                    worksheet.Cells["C" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["C" + row].Style.WrapText = true;

                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                    worksheet.Cells["D" + row].Value = item.Cantidad;
                    worksheet.Cells["D" + row].Style.Font.Size = 8;
                    worksheet.Cells["D" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Numberformat.Format = "#,##0";

                    worksheet.Cells["E" + row].Value = item.PreUnitario;
                    worksheet.Cells["E" + row].Style.Font.Size = 8;
                    worksheet.Cells["E" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["E" + row].Style.WrapText = true;

                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Numberformat.Format = "#,##0.00";



                    worksheet.Cells["F" + row].Value = item.PreTotal;
                    worksheet.Cells["F" + row].Style.Font.Size = 8;
                    worksheet.Cells["F" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["F" + row].Style.WrapText = true;

                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Numberformat.Format = "#,##0.00";


                    worksheet.Cells["G" + row].Value = item.Rnp;
                    worksheet.Cells["G" + row].Style.Font.Size = 8;
                    worksheet.Cells["G" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["G" + row].Style.WrapText = true;

                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["H" + row].Value = item.NRegistro;
                    worksheet.Cells["H" + row].Style.Font.Size = 8;
                    worksheet.Cells["H" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["I" + row].Value = item.VigProducto;
                    worksheet.Cells["I" + row].Style.Font.Size = 8;
                    worksheet.Cells["I" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["I" + row].Style.WrapText = true;

                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    worksheet.Cells["J" + row].Value = item.CertBPM;
                    worksheet.Cells["J" + row].Style.Font.Size = 8;
                    worksheet.Cells["J" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    worksheet.Cells["K" + row].Value = item.ProAnalisis;
                    worksheet.Cells["K" + row].Style.Font.Size = 8;
                    worksheet.Cells["K" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["L" + row].Value = item.ValiOferta;
                    worksheet.Cells["L" + row].Style.Font.Size = 8;
                    worksheet.Cells["L" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    worksheet.Cells["M" + row].Value = item.Marcapaisesprocedencia;
                    worksheet.Cells["M" + row].Style.Font.Size = 8;
                    worksheet.Cells["M" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Cells["N" + row].Value = item.CumpleEETT;
                    worksheet.Cells["N" + row].Style.Font.Size = 8;
                    worksheet.Cells["N" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["N" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    worksheet.Cells["O" + row].Value = item.PlazoEntrega;
                    worksheet.Cells["O" + row].Style.Font.Size = 8;
                    worksheet.Cells["O" + row].Style.Font.Name = "Arial";
                    worksheet.Cells["O" + row].Style.WrapText = true;
                    worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["O" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                    row++;
                }


                worksheet.Row(row).Height = 15.00;

                row++;
                worksheet.Cells["A" + row + ":G" + row].Merge = true;
                worksheet.Cells["A" + row + ":G" + row].Value = "* Declaro bajo juramento que mi representada cumple conlas especificaciones tecncias y requisitos de calificación.";
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row + ":G" + row].Style.WrapText = true;
                worksheet.Cells["A" + row + ":G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row + ":G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                row++;

                worksheet.Cells["A" + row ].Merge = true;
                worksheet.Cells["A" + row].Value = "NOTAS :";
                worksheet.Cells["A" + row].Style.Font.Bold =true;
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                row++;

                worksheet.Cells["A" + row + ":G" + row].Merge = true;
                worksheet.Cells["A" + row + ":G" + row].Value = "-En caso de que la presentación ofertada no se ajusta a la solicitada, deberá de indicar el número de unidades equivalentes. ";
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row + ":G" + row].Style.WrapText = true;
                worksheet.Cells["A" + row + ":G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row + ":G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;


                worksheet.Cells["A" + row + ":G" + row].Merge = true;
                worksheet.Cells["A" + row + ":G" + row].Value = "- Deberá adjuntar el protocolo de análisis y/o ficha técnica de cada producto cotizado.";
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row + ":G" + row].Style.WrapText = true;
                worksheet.Cells["A" + row + ":G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row + ":G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                row++;


                worksheet.Cells["A" + row +":G" + row].Merge = true;
                worksheet.Cells["A" + row + ":G" + row].Value = "- De no cumplir con la vigencia solicitada, el postor deberá  adjuntar a la cotización carta de compromiso de canje.";
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row + ":G" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row + ":G" + row].Style.WrapText = true;
                worksheet.Cells["A" + row + ":G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row + ":G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                 row++;
                 row++;
                 row++;
                worksheet.Cells["A" + row].Merge = true;
                worksheet.Cells["A" + row].Value = "Atentamente,";
                worksheet.Cells["A" + row].Style.Font.Size = 8;
                worksheet.Cells["A" + row].Style.Font.Name = "Arial";
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["A" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;



                worksheet.Cells["A" + (row+5) + ":B" + (row + 5)].Merge = true;
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Value = "FIRMA Y SELLO DEL REPRESENTANTE LEGAL (PROVEEDOR)";
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Style.Font.Size = 8;
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Style.Font.Name = "Arial";
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Style.WrapText = true;
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Style.Font.Bold = true;
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A" + (row + 5) + ":B" + (row + 5)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Merge = true;
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Value = "FIRMA Y SELLO DEL AREA USUARIA DEL INSN-SB";
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Style.Font.Size = 8;
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Style.Font.Name = "Arial";
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Style.Font.Bold = true;
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Style.WrapText = true;
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K" + (row + 5) + ":O" + (row + 5)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;



                ExcelPicture firmaCatizacion = worksheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(row-2, -8, 0, 85);
                firmaCatizacion.SetSize(265, 120);

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

            worksheet.Column(1).Width = 16.71 + 0.71;
            worksheet.Column(2).Width = 48.14 + 0.71;
            worksheet.Column(3).Width = 8.43 + 0.71;
            worksheet.Column(4).Width = 8.43 + 0.71;
            worksheet.Column(5).Width = 9.14 + 0.71;
            worksheet.Column(6).Width = 10.29 + 0.71;
            worksheet.Column(7).Width = 11.29 + 0.71;
            worksheet.Column(8).Width = 13.57 + 0.71;
            worksheet.Column(9).Width = 13.57 + 0.71;
            worksheet.Column(10).Width = 10.43 + 0.71;
            worksheet.Column(11).Width = 9.43 + 0.71;
            worksheet.Column(12).Width = 10.29 + 0.71;
            worksheet.Column(13).Width = 12 + 0.71;
            worksheet.Column(14).Width = 11.29 + 0.71;
            worksheet.Column(15).Width = 14.57 + 0.71;

            worksheet.Row(11).Height = 62.25;
            worksheet.Row(18).Height = 56.25;

        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["L4:O4"].Merge = true;
            worksheet.Cells["A5:M5"].Merge = true;
            worksheet.Cells["N5:O5"].Merge = true;
            worksheet.Cells["A6:M6"].Merge = true;
            worksheet.Cells["N6:O6"].Merge = true;
            worksheet.Cells["N7:O7"].Merge = true;
            worksheet.Cells["N8:O8"].Merge = true;
            worksheet.Cells["A11:O11"].Merge = true;
            worksheet.Cells["B12:L12"].Merge = true;
            worksheet.Cells["N12:O12"].Merge = true;
            worksheet.Cells["B13:O13"].Merge = true;
            worksheet.Cells["B14:O14"].Merge = true;
            worksheet.Cells["B15:O15"].Merge = true;
            worksheet.Cells["B16:O16"].Merge = true;
         
            worksheet.Cells["B17:F17"].Merge = true;
            worksheet.Cells["L17:M17"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A5:M5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N5:O5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N6:O6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N7:O7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N8:O8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B12:L12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["M12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N12:O12"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B13:O13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B14:O14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B15:O15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B16:O16"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

           

            worksheet.Cells["A17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B17:F17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["G17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["H17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["I17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["J17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["K17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["L17:M17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["O17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells["A18,B18,C18,D18,E18,F18,G18,H18,I18,J18,H18,I18,J8,K18,L18,M18,N18,O18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["N5:O5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
            worksheet.Cells["N6:O6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            worksheet.Cells["N7:O7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
            worksheet.Cells["N8:O8"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            worksheet.Cells["A12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
            worksheet.Cells["M12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
            worksheet.Cells["A13"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
            worksheet.Cells["A14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
            worksheet.Cells["A15"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
            worksheet.Cells["A16"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));

         
            worksheet.Cells["A17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
            worksheet.Cells["G17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
            worksheet.Cells["H17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
            worksheet.Cells["J17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
            worksheet.Cells["K17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
            worksheet.Cells["N17"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

            worksheet.Cells["A18:O18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#808080"));
        }

        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
