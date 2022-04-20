using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public static class Formato61_Report
    {
        public static string Exportar(Image firma, Image logoUnilene, Formato61_Model cotizacion)
        {
            byte[] file;
            string reporte = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("EsSalud Incor");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 1, 0, 1);
                imagenUnilene.SetSize(190, 53);

                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);

                workSheet.Cells["B3"].Value = "FORMATO DE COTIZACION DE BIENES";
                workSheet.Cells["B3"].Style.Font.Size = 12;
                workSheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                workSheet.Cells["I5"].Value = "Cotizacion N° " + cotizacion.Nro_Cotizacion + " -Unilene SAC";
                workSheet.Cells["I5"].Style.Font.Size = 10;
                workSheet.Cells["I5"].Style.Font.UnderLine = true;

                workSheet.Cells["B7"].Value = "ESSALUD INCOR";
                workSheet.Cells["B7"].Style.Font.Size = 11;

                workSheet.Cells["B8"].Value = "Presente.- ";
                workSheet.Cells["B8"].Style.Font.Size = 11;

                workSheet.Cells["B10"].Value = "De mi consideración";
                workSheet.Cells["B10"].Style.Font.Size = 11;

                workSheet.Cells["B8"].Value = "De mi consideración";
                workSheet.Cells["B8"].Style.Font.Size = 11;

                workSheet.Cells["B12"].Value = "En respuesta la solicitud de cotización sobre la “ADQUISICIÓN  DE MATERIAL MEDICO..”, Y " +
                    "después de haber \nanalizado las especificaciones técnicas del mencionado requerimiento, las mismas que acepto en todos sus " +
                    "extremos, indico que cumplo con TODOS \nlos requerimientos solicitados.";
                workSheet.Cells["B12"].Style.Font.Size = 8;
                workSheet.Cells["B12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B12"].Style.WrapText = true;

                workSheet.Cells["B13"].Value = "Asimismo, declaro que las características técnicas de los bienes cotizados por mi representada se ajustan " +
                    "a lo requerido \npor su ENTIDAD.En tal sentido, indico que el costo total por lo requerido es lo que detallo a continuación: ";
                workSheet.Cells["B13"].Style.Font.Size = 8;
                workSheet.Cells["B13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B13"].Style.WrapText = true;


                workSheet.Cells["B14"].Value = "ITEM";
                workSheet.Cells["B14"].Style.Font.Size = 8;
                workSheet.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["C14"].Value = "CÓDIGO\nSAP";
                workSheet.Cells["C14"].Style.Font.Size = 8;
                workSheet.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C14"].Style.WrapText = true;

                workSheet.Cells["D14"].Value = "DESCRIPCIÓN";
                workSheet.Cells["D14"].Style.Font.Size = 8;
                workSheet.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["E14"].Value = "CANTIDAD";
                workSheet.Cells["E14"].Style.Font.Size = 8;
                workSheet.Cells["E14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["F14"].Value = "UM";
                workSheet.Cells["F14"].Style.Font.Size = 8;
                workSheet.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["G14"].Value = "VIGENCIA DEL\nPRODUCTO";
                workSheet.Cells["G14"].Style.Font.Size = 8;
                workSheet.Cells["G14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G14"].Style.WrapText = true;

                workSheet.Cells["H14"].Value = "MARCA";
                workSheet.Cells["H14"].Style.Font.Size = 8;
                workSheet.Cells["H14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["I14"].Value = "MODELO";
                workSheet.Cells["I14"].Style.Font.Size = 8;
                workSheet.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["J14"].Value = "PAIS DE\nPROCEDENCIA";
                workSheet.Cells["J14"].Style.Font.Size = 8;
                workSheet.Cells["J14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J14"].Style.WrapText = true;

                workSheet.Cells["K14"].Value = "OBSERVACIONES";
                workSheet.Cells["K14"].Style.Font.Size = 8;
                workSheet.Cells["K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["L14"].Value = "PRECION\nUNITARIO";
                workSheet.Cells["L14"].Style.Font.Size = 8;
                workSheet.Cells["L14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L14"].Style.WrapText = true;

                workSheet.Cells["M14"].Value = "PRECIO TOTAL";
                workSheet.Cells["M14"].Style.Font.Size = 8;
                workSheet.Cells["M14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 15;

                foreach (Formato61_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Cells["B" + row].Value = row - 14;
                    workSheet.Cells["B" + row].Style.Font.Size = 12;
                    workSheet.Cells["B" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + row].Value = item.Sap;
                    workSheet.Cells["C" + row].Style.Font.Size = 10;
                    workSheet.Cells["C" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + row].Value = item.Descripcion;
                    workSheet.Cells["D" + row].Style.Font.Size = 12;
                    workSheet.Cells["D" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + row].Style.WrapText = true;
                    workSheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + row].Value = item.Cantidad;
                    workSheet.Cells["E" + row].Style.Font.Size = 10;
                    workSheet.Cells["E" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + row].Style.Numberformat.Format = "#,##0";
                    workSheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + row].Value = item.Um;
                    workSheet.Cells["F" + row].Style.Font.Size = 8;
                    workSheet.Cells["F" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + row].Value = item.Vigencia;
                    workSheet.Cells["G" + row].Style.Font.Size = 12;
                    workSheet.Cells["G" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + row].Style.WrapText = true;
                    workSheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + row].Value = item.Marca;
                    workSheet.Cells["H" + row].Style.Font.Size = 10;
                    workSheet.Cells["H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + row].Style.WrapText = true;
                    workSheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + row].Value = item.Modelo;
                    workSheet.Cells["I" + row].Style.Font.Size = 10;
                    workSheet.Cells["I" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + row].Style.WrapText = true;
                    workSheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + row].Value = item.Procedencia;
                    workSheet.Cells["J" + row].Style.Font.Size = 12;
                    workSheet.Cells["J" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + row].Style.WrapText = true;
                    workSheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + row].Value = item.Observaciones;
                    workSheet.Cells["K" + row].Style.Font.Size = 12;
                    workSheet.Cells["K" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["k" + row].Style.WrapText = true;
                    workSheet.Cells["k" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + row].Value = item.PrecioUnitario;
                    workSheet.Cells["L" + row].Style.Font.Size = 12;
                    workSheet.Cells["L" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + row].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["L" + row].Style.WrapText = true;
                    workSheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + row].Value = item.PrecioTotal;
                    workSheet.Cells["M" + row].Style.Font.Size = 12;
                    workSheet.Cells["M" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    workSheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.00";
                    workSheet.Cells["M" + row].Style.WrapText = true;
                    workSheet.Cells["M" + row].Style.Font.Bold = true;
                    workSheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    row++;
                }

                workSheet.Cells["B" + row + ":L" + row].Merge = true;
                workSheet.Cells["B" + row + ":L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["B" + row].Value = "TOTAL S/.";
                workSheet.Cells["B" + row].Style.Font.Size = 12;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                workSheet.Cells["M" + row].Value = cotizacion.Monto_Total;
                workSheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.00";
                workSheet.Cells["M" + row].Style.Font.Size = 12;
                workSheet.Cells["M" + row].Style.Font.Bold = true;
                workSheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;

                workSheet.Cells["B" + row + ":M" + row].Merge = true;
                workSheet.Cells["B" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.Font.Size = 8;
                workSheet.Cells["B" + row].Style.WrapText = true;
                workSheet.Row(row).Height = 48.75;
                workSheet.Cells["B" + row].Value = "La propuesta se emite considerando todas las condiciones señaladas en el requerimiento e incluye " +
                    "todos los tributos, seguros, transporte, inspecciones, pruebas y, de ser el caso, los costos laborales conforme la legislación vigente, " +
                    "así como cualquier otro concepto que pueda tener incidencia sobre el costo del bien y/o servicio a contratar excepto la de aquellos " +
                    "proveedores que gocen de alguna exoneración legal, no incluirán en el precio de oferta los tributos respetivos. Así mismo, declaro bajo " +
                    "juramento que, mi persona y/o mi representada no cuenta con impedimentos para contratar con el Estado, confirme lo establece el artículo 11 del " +
                    "Texto Único ordenado de la Ley N°30225, Ley de Contrataciones del Estado, aprobado por Decreto Supremos N°082-2019-EF.";

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "Razón Social";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_RazonSocial;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "N° R.U.C.";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_Ruc;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "Plazo de Entrega (días calendario)";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_PlazaEntrega;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "Forma de pago";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_FormaPago;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 10;
                workSheet.Cells["B" + row].Value = "Correo Electrónico";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 11;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_Correo;
                workSheet.Cells["E" + row].Style.Font.UnderLine = true;
                workSheet.Cells["E" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#0000FF"));

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 10;
                workSheet.Cells["B" + row].Value = "Teléfono Fijo";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_Fijo;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "Persona de Contacto";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_Contacto;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "Teléfono Móvil";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_Movil;

                row++;

                workSheet.Cells["B" + row + ":D" + row].Merge = true;
                workSheet.Cells["B" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Style.Font.Bold = true;
                workSheet.Cells["B" + row].Style.Font.Size = 11;
                workSheet.Cells["B" + row].Value = "Vigencia de Oferta (días calendario)";

                workSheet.Cells["E" + row + ":M" + row].Merge = true;
                workSheet.Cells["E" + row + ":M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E" + row].Style.Font.Size = 10;
                workSheet.Cells["E" + row].Value = cotizacion.Prov_Vigencia;

                row++;

                ExcelPicture firmaCatizacion = workSheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(row, -10, 6, 60);
                firmaCatizacion.SetSize(130, 85);

                row += 5;

                workSheet.Cells["B" + row + ":M" + row].Merge = true;
                workSheet.Cells["B" + row + ":M" + (row + 1)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                workSheet.Cells["B" + row].Style.Font.Size = 8;
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Value = "Firma, nombres y apellidos del proveedor o representante legal";

                row++;

                workSheet.Cells["B" + row + ":M" + row].Merge = true;
                workSheet.Cells["B" + row].Style.Font.Size = 8;
                workSheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B" + row].Value = "o persona autorizada para emitir cotizaciones ";


                TextoNegrita(workSheet);

                workSheet.View.ZoomScale = 150;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

            }

            return reporte;
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 0.25;
            worksheet.Column(2).Width = 6.57 + 0.71;
            worksheet.Column(3).Width = 10.71 + 0.71;
            worksheet.Column(4).Width = 25.14 + 0.71;
            worksheet.Column(5).Width = 7.71 + 0.71;
            worksheet.Column(6).Width = 3.71 + 0.71;
            worksheet.Column(7).Width = 14.86 + 0.71;
            worksheet.Column(8).Width = 12.71 + 0.71;
            worksheet.Column(9).Width = 13.86 + 0.71;
            worksheet.Column(10).Width = 10.71 + 0.71;
            worksheet.Column(11).Width = 11.57 + 0.71;
            worksheet.Column(12).Width = 8.57 + 0.71;
            worksheet.Column(13).Width = 17.14 + 0.71;

            worksheet.Row(4).Height = 8.25;
            worksheet.Row(5).Height = 13.5;
            worksheet.Row(6).Height = 11.25;
            worksheet.Row(7).Height = 13.5;
            worksheet.Row(8).Height = 11.25;
            worksheet.Row(9).Height = 0;
            worksheet.Row(10).Height = 15;
            worksheet.Row(11).Height = 3.75;
            worksheet.Row(12).Height = 31.5;
            worksheet.Row(13).Height = 29.25;
            worksheet.Row(14).Height = 39.75;

        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["B3:M3, I5:K5"].Merge = true;
            worksheet.Cells["B12:M12, B13:M13"].Merge = true;
        }

        private static void PintarCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["B14, C14, D14, E14, F14, G14, H14, I14, J14, K14, L14, M14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["B14, C14, D14, E14, F14, G14, H14, I14, J14, K14, L14, M14"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
        }

        private static void BordesCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["B14, C14, D14, E14, F14, G14, H14, I14, J14, K14, L14, M14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
            worksheet.Cells["B3, I5, B7"].Style.Font.Bold = true;
            worksheet.Cells["B14, C14, D14, E14, F14, G14, H14, I14, J14, K14, L14, M14"].Style.Font.Bold = true;
        }
    }
}
