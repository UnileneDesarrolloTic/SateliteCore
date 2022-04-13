using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato18_Report
    {
        public static string Exportar(Image logoUnilene, Coti_Formato18_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Cotización LoretoF1");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 0);
                imagenUnilene.SetSize(120, 30);

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                workSheet.Cells["T1"].Value = "Fecha:";
                workSheet.Cells["T1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["T1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["T1"].Style.Font.Size = 10;

                workSheet.Cells["U1"].Value = cotizacion.Prov_Fecha;
                workSheet.Cells["U1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["U1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["U1"].Style.Font.Size = 10;
                workSheet.Cells["U1"].Style.Numberformat.Format = "dd/MM/yyyy";


                workSheet.Cells["A3"].Value = "Cotiz Nº " + cotizacion.Prov_NroCotizacion + "2022-Unilene S.A.C";
                workSheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A3"].Style.Font.Bold = true;
                workSheet.Cells["A3"].Style.Font.Size = 10;

                workSheet.Cells["A4"].Value = "OFICINAS: JR. NAPO 450 \n" +
                    "CENTRAL TELEFONICA: 7487006 / 7487000 \n" +
                    "contactenos@unilene.com \n" +
                    "www.unilene.com contactenosunilene.com";
                workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A4"].Style.Font.Size = 10;
                workSheet.Cells["A4"].Style.WrapText = true;

                workSheet.Cells["A6"].Value = "Codigo";
                workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A6"].Style.Font.Bold = true;
                workSheet.Cells["A6"].Style.Font.Size = 10;

                workSheet.Cells["C6"].Value = cotizacion.Prov_Codigo;
                workSheet.Cells["C6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C6"].Style.Font.Size = 10;

                workSheet.Cells["J6"].Value = "Asesor Comercial:";
                workSheet.Cells["J6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J6"].Style.Font.Bold = true;
                workSheet.Cells["J6"].Style.Font.Size = 10;

                workSheet.Cells["P6"].Value = cotizacion.Prov_AsesComercial;
                workSheet.Cells["P6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P6"].Style.Font.Size = 10;

                workSheet.Cells["A7"].Value = "Señor(es):";
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A7"].Style.Font.Bold = true;
                workSheet.Cells["A7"].Style.Font.Size = 10;

                workSheet.Cells["C7"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C7"].Style.Font.Size = 10;

                workSheet.Cells["J7"].Value = "Email:";
                workSheet.Cells["J7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J7"].Style.Font.Bold = true;
                workSheet.Cells["J7"].Style.Font.Size = 10;

                workSheet.Cells["P7"].Value = cotizacion.Prov_Email;
                workSheet.Cells["P7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P7"].Style.Font.Size = 10;


                workSheet.Cells["A8"].Value = "RUC:";
                workSheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A8"].Style.Font.Bold = true;
                workSheet.Cells["A8"].Style.Font.Size = 10;

                workSheet.Cells["C8"].Value = cotizacion.Prov_Ruc1;
                workSheet.Cells["C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C8"].Style.Font.Size = 10;

                workSheet.Cells["J8"].Value = "Telefono:";
                workSheet.Cells["J8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J8"].Style.Font.Bold = true;
                workSheet.Cells["J8"].Style.Font.Size = 10;

                workSheet.Cells["P8"].Value = cotizacion.Prov_Telefono2;
                workSheet.Cells["P8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P8"].Style.Font.Size = 10;


                workSheet.Cells["A9"].Value = "Dirección:";
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A9"].Style.Font.Bold = true;
                workSheet.Cells["A9"].Style.Font.Size = 10;

                workSheet.Cells["C9"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["C9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C9"].Style.Font.Size = 10;

                workSheet.Cells["J9"].Value = "Validez de la Oferta:";
                workSheet.Cells["J9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J9"].Style.Font.Bold = true;
                workSheet.Cells["J9"].Style.Font.Size = 10;

                workSheet.Cells["P9"].Value = cotizacion.Prov_VigOferta;
                workSheet.Cells["P9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P9"].Style.Font.Size = 10;


                workSheet.Cells["A10"].Value = "Teléfono:";
                workSheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A10"].Style.Font.Bold = true;
                workSheet.Cells["A10"].Style.Font.Size = 10;

                workSheet.Cells["C10"].Value = cotizacion.Prov_Telefono1;
                workSheet.Cells["C10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C10"].Style.Font.Size = 10;

                workSheet.Cells["J10"].Value = "Plazo de Entrega:";
                workSheet.Cells["J10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J10"].Style.Font.Bold = true;
                workSheet.Cells["J10"].Style.Font.Size = 10;

                workSheet.Cells["P10"].Value = cotizacion.Prov_PlazoEntrega;
                workSheet.Cells["P10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P10"].Style.Font.Size = 10;


                workSheet.Cells["A12"].Value = "Fax:";
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A12"].Style.Font.Bold = true;
                workSheet.Cells["A12"].Style.Font.Size = 10;

                workSheet.Cells["C12"].Value = cotizacion.Prov_Fax;
                workSheet.Cells["C12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C12"].Style.Font.Size = 10;

                workSheet.Cells["J12"].Value = "Lugar de Entrega:";
                workSheet.Cells["J12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J12"].Style.Font.Bold = true;
                workSheet.Cells["J12"].Style.Font.Size = 10;

                workSheet.Cells["P12"].Value = cotizacion.Prov_LugarEntrega;
                workSheet.Cells["P12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P12"].Style.Font.Size = 10;


                workSheet.Cells["A13"].Value = "Atención:";
                workSheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A13"].Style.Font.Bold = true;
                workSheet.Cells["A13"].Style.Font.Size = 10;

                workSheet.Cells["C13"].Value = cotizacion.Prov_Atencion;
                workSheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C13"].Style.Font.Size = 10;

                workSheet.Cells["J13"].Value = "Garantía:";
                workSheet.Cells["J13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J13"].Style.Font.Bold = true;
                workSheet.Cells["J13"].Style.Font.Size = 10;


                workSheet.Cells["P13"].Value = cotizacion.Prov_Garantia;
                workSheet.Cells["P13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P13"].Style.Font.Size = 10;



                workSheet.Cells["A14"].Value = "Cond. Venta:";
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A14"].Style.Font.Bold = true;
                workSheet.Cells["A14"].Style.Font.Size = 10;

                workSheet.Cells["C14"].Value = cotizacion.Prov_CondVenta;
                workSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C14"].Style.Font.Size = 10;

                workSheet.Cells["J14"].Value = "Vigencia Producto:";
                workSheet.Cells["J14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["J14"].Style.Font.Bold = true;
                workSheet.Cells["J14"].Style.Font.Size = 10;

                workSheet.Cells["P14"].Value = cotizacion.Prov_vigProducto;
                workSheet.Cells["P14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P14"].Style.Font.Size = 10;
                workSheet.Cells["P14"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                workSheet.Cells["A15"].Value = "Fecha:";
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A15"].Style.Font.Bold = true;
                workSheet.Cells["A15"].Style.Font.Size = 10;

                workSheet.Cells["C15"].Value = cotizacion.Prov_Fecha;
                workSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["C15"].Style.Font.Size = 10;

                workSheet.Cells["J15"].Value = "Ref:";
                workSheet.Cells["J15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["J15"].Style.Font.Bold = true;
                workSheet.Cells["J15"].Style.Font.Size = 10;

                workSheet.Cells["P15"].Value = cotizacion.Prov_Refencia;
                workSheet.Cells["P15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["P15"].Style.Font.Size = 10;


                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);

                //DETALLE

                workSheet.Cells["A17"].Value = "COD.SUT";
                workSheet.Cells["A17"].Style.Font.Size = 8;
                workSheet.Cells["A17"].Style.Font.Name = "Arial";
                workSheet.Cells["A17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A17"].Style.WrapText = true;


                workSheet.Cells["B17"].Value = "DESCRIPCION";
                workSheet.Cells["B17"].Style.Font.Size = 8;
                workSheet.Cells["B17"].Style.Font.Name = "Arial";
                workSheet.Cells["B17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["B17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B17"].Style.WrapText = true;
                workSheet.Cells["B17"].Style.Font.Bold = true;

                workSheet.Cells["F17"].Value = "MARCA";
                workSheet.Cells["F17"].Style.Font.Size = 8;
                workSheet.Cells["F17"].Style.Font.Name = "Arial";
                workSheet.Cells["F17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["F17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F17"].Style.WrapText = true;
                workSheet.Cells["F17"].Style.Font.Bold = true;


                workSheet.Cells["H17"].Value = "PRESENTACION";
                workSheet.Cells["H17"].Style.Font.Size = 8;
                workSheet.Cells["H17"].Style.Font.Name = "Arial";
                workSheet.Cells["H17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["H17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H17"].Style.WrapText = true;
                workSheet.Cells["H17"].Style.Font.Bold = true;

                workSheet.Cells["M17"].Value = "PROCED.";
                workSheet.Cells["M17"].Style.Font.Size = 8;
                workSheet.Cells["M17"].Style.Font.Name = "Arial";
                workSheet.Cells["M17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["M17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M17"].Style.WrapText = true;
                workSheet.Cells["M17"].Style.Font.Bold = true;


                workSheet.Cells["N17"].Value = "UND MED";
                workSheet.Cells["N17"].Style.Font.Size = 8;
                workSheet.Cells["N17"].Style.Font.Name = "Arial";
                workSheet.Cells["N17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["N17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N17"].Style.WrapText = true;
                workSheet.Cells["N17"].Style.Font.Bold = true;

                workSheet.Cells["Q17"].Value = "CANT.UNITARIA";
                workSheet.Cells["Q17"].Style.Font.Size = 8;
                workSheet.Cells["Q17"].Style.Font.Name = "Arial";
                workSheet.Cells["Q17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["Q17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q17"].Style.WrapText = true;
                workSheet.Cells["Q17"].Style.Font.Bold = true;

                workSheet.Cells["R17"].Value = "PLAZO ENTREGA";
                workSheet.Cells["R17"].Style.Font.Size = 8;
                workSheet.Cells["R17"].Style.Font.Name = "Arial";
                workSheet.Cells["R17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["R17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R17"].Style.WrapText = true;
                workSheet.Cells["R17"].Style.Font.Bold = true;

                workSheet.Cells["S17"].Value = "PRECIO UNITARIO";
                workSheet.Cells["S17"].Style.Font.Size = 8;
                workSheet.Cells["S17"].Style.Font.Name = "Arial";
                workSheet.Cells["S17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["S17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S17"].Style.WrapText = true;
                workSheet.Cells["S17"].Style.Font.Bold = true;

                workSheet.Cells["V17"].Value = "TOTAL";
                workSheet.Cells["V17"].Style.Font.Size = 8;
                workSheet.Cells["V17"].Style.Font.Name = "Arial";
                workSheet.Cells["V17"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["V17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["V17"].Style.WrapText = true;
                workSheet.Cells["V17"].Style.Font.Bold = true;



                int index = 18;

                foreach (Coti_Formato18_Detalle item in cotizacion.Detalle)
                {

                   
                    workSheet.Row(index).Height = 25.5;
                    workSheet.Cells["A" + index].Value = item.CodSut;
                    workSheet.Cells["A" + index].Style.Font.Size = 10;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.WrapText = true;

                    workSheet.Cells["B" + index + ":E" + index].Value = item.Descripcion;
                    workSheet.Cells["B" + index + ":E" + index].Style.Font.Size = 10;
                    workSheet.Cells["B" + index + ":E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index + ":E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["B" + index + ":E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index + ":E" + index].Style.WrapText = true;
                    workSheet.Cells["B" + index + ":E" + index].Merge = true;
                  

                    workSheet.Cells["F" + index + ":G" + index].Value = item.Marca;
                    workSheet.Cells["F" + index + ":G" + index].Style.Font.Size = 10;
                    workSheet.Cells["F" + index + ":G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index + ":G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["F" + index + ":G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["F" + index + ":G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index + ":G" + index].Style.WrapText = true;
                    workSheet.Cells["F" + index + ":G" + index].Merge = true;

                    workSheet.Cells["H" + index + ":L" + index].Value = item.Presentacion;
                    workSheet.Cells["H" + index + ":L" + index].Style.Font.Size = 10;
                    workSheet.Cells["H" + index + ":L" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index + ":L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["H" + index + ":L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["H" + index + ":L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index + ":L" + index].Style.WrapText = true;
                    workSheet.Cells["H" + index + ":L" + index].Merge = true;


                    workSheet.Cells["M" + index].Value = item.Procedencia;
                    workSheet.Cells["M" + index].Style.Font.Size = 10;
                    workSheet.Cells["M" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;

                    workSheet.Cells["N" + index + ":P" + index].Value = item.UnidMedida;
                    workSheet.Cells["N" + index + ":P" + index].Style.Font.Size = 10;
                    workSheet.Cells["N" + index + ":P" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["N" + index + ":P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["N" + index + ":P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index + ":P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["N" + index + ":P" + index].Style.WrapText = true;
                    workSheet.Cells["N" + index + ":P" + index].Merge = true;


                    workSheet.Cells["Q" + index].Value = item.Cantidad;
                    workSheet.Cells["Q" + index].Style.Font.Size = 10;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;
                    workSheet.Cells["Q" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["R" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["R" + index].Style.Font.Size = 10;
                    workSheet.Cells["R" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["R" + index].Style.WrapText = true;



                    workSheet.Cells["S" + index + ":U" + index].Value = item.PreUnitario;
                    workSheet.Cells["S" + index + ":U" + index].Style.Font.Size = 10;
                    workSheet.Cells["S" + index + ":U" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["S" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["S" + index + ":U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["S" + index + ":U" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    workSheet.Cells["S" + index + ":U" + index].Style.WrapText = true;
                    workSheet.Cells["S" + index + ":U" + index].Merge = true;
                    workSheet.Cells["S" + index + ":U" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["V" + index + ":W" + index].Value = item.PreTotal;
                    workSheet.Cells["V" + index + ":W" + index].Style.Font.Size = 10;
                    workSheet.Cells["V" + index + ":W" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["V" + index + ":W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells["V" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["V" + index + ":W" + index].Style.WrapText = true;
                    workSheet.Cells["V" + index + ":W" + index].Merge = true;
                    workSheet.Cells["V" + index + ":W" + index].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["V" + index + ":W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;


                    index++;
                }

                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":U" + index].Value = "Monto Total     S/.";
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index + ":U" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":U" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells["A" + index + ":U" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["A" + index + ":U" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":U" + index].Merge = true;
                workSheet.Cells["A" + index + ":U" + index].Style.Numberformat.Format = "#,##0.00";


                workSheet.Cells["V" + index + ":W" + index].Value = cotizacion.Prov_ValorTotal;
                workSheet.Cells["V" + index + ":W" + index].Style.Font.Size = 10;
                workSheet.Cells["V" + index + ":W" + index].Style.Font.Name = "Arial";
                workSheet.Cells["V" + index + ":W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["V" + index + ":W" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["V" + index + ":W" + index].Style.WrapText = true;
                workSheet.Cells["V" + index + ":W" + index].Merge = true;
                workSheet.Cells["V" + index + ":W" + index].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["V" + index + ":W" + index].Style.Font.Bold = true;



                index++;
                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":J" + index].Value = "CONDICIONES GENERALES :";
                workSheet.Cells["A" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":J" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index + ":J" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":D" + index].Value = "PRECIO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;

                workSheet.Cells["E" + index + ":J" + index].Value = "EN Soles INCLUIDO I.G.V";
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["E" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["E" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["E" + index + ":J" + index].Merge = true;


                index++;
                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":D" + index].Value = "FORMA DE PAGO";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;

                workSheet.Cells["E" + index + ":J" + index].Value =cotizacion.Prov_FormaPago;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["E" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["E" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["E" + index + ":J" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":D" + index].Value = "LABORATORIO FABRICANTE";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;

                workSheet.Cells["E" + index + ":J" + index].Value = cotizacion.Prov_laboratorio;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["E" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["E" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["E" + index + ":J" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":D" + index].Value = "RUC";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;

                workSheet.Cells["E" + index + ":J" + index].Value = cotizacion.Prov_Ruc2;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["E" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["E" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["E" + index + ":J" + index].Merge = true;

                index++;
                workSheet.Row(index).Height = 17.00;
                workSheet.Cells["A" + index + ":D" + index].Value = "FECHA DE VENCIMIENTO DE COTIZACION";
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":D" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":D" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":D" + index].Merge = true;

                workSheet.Cells["E" + index + ":J" + index].Value = cotizacion.Prov_FechaVenci;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["E" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["E" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["E" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["E" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["E" + index + ":J" + index].Merge = true;
                workSheet.Cells["E" + index + ":J" + index].Style.Numberformat.Format = "dd/MM/yyyy";

                index++;
                workSheet.Row(index).Height = 26.00;
                workSheet.Cells["A" + index + ":J" + index].Value = "Sin otro particular y a la espera de sus gratas ordenes, quedamos de Ustedes.";
                workSheet.Cells["A" + index + ":J" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":J" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":J" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":J" + index].Merge = true;


                index++;
                workSheet.Row(index).Height = 102.00;


                index++;
                workSheet.Row(index).Height = 36.5;
                workSheet.Cells["A" + index + ":W" + index].Value = "NOTA : Vencida la cotización comunicarse con el Area Comercial de nuestra empresa a los teléfonos señalados a fin de actualizar la información vía E-mail indicando"
                        +"la nueva fecha de entrega de nuestros productos y así poder atenderlos evitando retrasos innecesarios"+
                        "UNILENE NO SE HACE RESPONSABLE SI VENCIDA LA COTIZACIÓN REMITEN ORDENES DE COMPRA SIN COORDINACION PREVIA DE LA ENTREGA";
                workSheet.Cells["A" + index + ":W" + index].Style.Font.Size = 8;
                workSheet.Cells["A" + index + ":W" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index + ":W" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index + ":W" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells["A" + index + ":W" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":W" + index].Merge = true;




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

            workSheet.Column(1).Width = 9.71 + 0.71;
            workSheet.Column(2).Width = 0.5 + 0.71;
            workSheet.Column(3).Width = 8.71 + 0.71;
            workSheet.Column(4).Width = 11 + 0.71;
            workSheet.Column(5).Width = 6 + 0.71;

            workSheet.Column(6).Width = 8.57 + 0.71;
            workSheet.Column(7).Width = 3.21 + 0.71;
            workSheet.Column(8).Width = 0.0;
            workSheet.Column(9).Width = 6.29 + 0.71;
            workSheet.Column(10).Width = 1.14 + 0.71;

            workSheet.Column(11).Width = 0.0 + 0.71;
            workSheet.Column(12).Width = 2.71 + 0.71;
            workSheet.Column(13).Width = 9.0 + 0.71;
            workSheet.Column(14).Width = 1.43 + 0.71;
            workSheet.Column(15).Width = 0.5 + 0.71;

            workSheet.Column(16).Width = 1.57 + 0.71;
            workSheet.Column(17).Width = 9.57 + 0.71;
            workSheet.Column(18).Width = 9.43 + 0.71;
            workSheet.Column(19).Width = 0.75 + 0.71;
            workSheet.Column(20).Width = 5.71 + 0.71;

            workSheet.Column(21).Width = 2.29 + 0.71;
            workSheet.Column(22).Width = 6.00 + 0.71;
            workSheet.Column(23).Width = 3.86 + 0.71;


            workSheet.Row(1).Height = 17.00;
            workSheet.Row(2).Height = 6.25;
            workSheet.Row(3).Height = 17.00;

            workSheet.Row(4).Height = 48.75;

            workSheet.Row(5).Height = 5;
            workSheet.Row(6).Height = 12.75;
            workSheet.Row(7).Height = 12.75;
            workSheet.Row(8).Height = 12.75;
            workSheet.Row(9).Height = 12.75;
            workSheet.Row(10).Height = 12.75;
            workSheet.Row(11).Height = 0.0;
            workSheet.Row(12).Height = 12.75;
            workSheet.Row(13).Height = 12.75;
            workSheet.Row(14).Height = 22.5;
            workSheet.Row(15).Height = 12.75;
            workSheet.Row(16).Height = 19.25;
            workSheet.Row(17).Height = 29.00;


        }
        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["U1:W1"].Merge = true;
            workSheet.Cells["A3:W3"].Merge = true;
            workSheet.Cells["A4:G4"].Merge = true;
            workSheet.Cells["C6:G6"].Merge = true;
            workSheet.Cells["J6:N6"].Merge = true;
            workSheet.Cells["P6:V6"].Merge = true;

            workSheet.Cells["C7:G7"].Merge = true;
            workSheet.Cells["J7:N7"].Merge = true;
            workSheet.Cells["P7:V7"].Merge = true;

            workSheet.Cells["C8:G8"].Merge = true;
            workSheet.Cells["J8:N8"].Merge = true;
            workSheet.Cells["P8:V8"].Merge = true;

            workSheet.Cells["C9:G9"].Merge = true;
            workSheet.Cells["J9:N9"].Merge = true;
            workSheet.Cells["P9:V9"].Merge = true;

            workSheet.Cells["C9:G9"].Merge = true;
            workSheet.Cells["J9:N9"].Merge = true;
            workSheet.Cells["P9:V9"].Merge = true;

            workSheet.Cells["C10:G10"].Merge = true;
            workSheet.Cells["J10:N10"].Merge = true;
            workSheet.Cells["P10:V10"].Merge = true;

            workSheet.Cells["C12:G12"].Merge = true;
            workSheet.Cells["J12:N12"].Merge = true;
            workSheet.Cells["P12:V12"].Merge = true;

            workSheet.Cells["C12:G12"].Merge = true;
            workSheet.Cells["J12:N12"].Merge = true;
            workSheet.Cells["P12:V12"].Merge = true;

            workSheet.Cells["C15:G15"].Merge = true;
            workSheet.Cells["J15:N15"].Merge = true;
            workSheet.Cells["P15:V15"].Merge = true;

            workSheet.Cells["C15:G15"].Merge = true;
            workSheet.Cells["J15:N15"].Merge = true;
            workSheet.Cells["P15:V15"].Merge = true;

            //DETALLE
            workSheet.Cells["B17:E17"].Merge = true;
            workSheet.Cells["F17:G17"].Merge = true;
            workSheet.Cells["H17:L17"].Merge = true;
            workSheet.Cells["N17:P17"].Merge = true;
            workSheet.Cells["S17:U17"].Merge = true;
            workSheet.Cells["V17:W17"].Merge = true;
        }
        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B17:E17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F17:G17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H17:L17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N17:P17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S17:U17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["V17:W17"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {

        }
    }
}
