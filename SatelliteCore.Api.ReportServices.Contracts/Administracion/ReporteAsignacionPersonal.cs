using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.AsignacionPersonal;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.HistorialPeriodo;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Administracion
{
    public class ReporteAsignacionPersonal
    {
        public string GenerarReporte(IEnumerable<DatosFormatoPersonaAsignacionExportModel> datos, IEnumerable<DatosFormatoListadoPersonalAsignacion> datosListadoPersonal, bool reporteAsistencia, bool listadoPersonal  )
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                if(reporteAsistencia == true)
                    CuerpoAsignacionPersonal(excelPackage, datos);
                
                if(listadoPersonal == true)
                    CuerpoListadoPersonal(excelPackage, datosListadoPersonal);



                    file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);


                return reporte;

            }

        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 10.86 + 2.71;
            worksheet.Column(2).Width = 35.86 + 2.71;
            worksheet.Column(3).Width = 20.43 + 2.71;
            worksheet.Column(4).Width = 20.71 + 2.71;
            worksheet.Column(5).Width = 20.14 + 2.71;
            worksheet.Column(6).Width = 10.14 + 2.71;
            worksheet.Column(7).Width = 10.14 + 2.71;


            //worksheet.Row(4).Height = 14.25;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:G3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            worksheet.Cells["A3:G3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#07438d"));

        }

        private static void CuerpoAsignacionPersonal(ExcelPackage excelPackage, IEnumerable<DatosFormatoPersonaAsignacionExportModel> datos)
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Reporte Asignacion Personal");
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

            ConfigurarTamanioDeCeldas(worksheet);
            pintarCabecera(worksheet);

            worksheet.Cells["A1:D1"].Merge = true;
            worksheet.Cells["A1:D1"].Value = "INFORMACIÓN ASIGNACIÓN DE PERSONAL";
            worksheet.Cells["A1:D1"].Style.Font.Size = 16;
            worksheet.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["A3"].Value = "CODIGO";
            worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A3"].Style.Font.Size = 10;
            worksheet.Cells["A3"].Style.WrapText = true;
            worksheet.Cells["A3"].Style.Font.Bold = true;
            worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["B3"].Value = "NOMBRE DEL TRABAJADOR";
            worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B3"].Style.Font.Size = 10;
            worksheet.Cells["B3"].Style.WrapText = true;
            worksheet.Cells["B3"].Style.Font.Bold = true;
            worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["C3"].Value = "NOMBRE AREA ";
            worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["C3"].Style.Font.Size = 10;
            worksheet.Cells["C3"].Style.WrapText = true;
            worksheet.Cells["C3"].Style.Font.Bold = true;
            worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["D3"].Value = "FECHA ASIGNACION ";
            worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["D3"].Style.Font.Size = 10;
            worksheet.Cells["D3"].Style.Font.Bold = true;
            worksheet.Cells["D3"].Style.WrapText = true;
            worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["E3"].Value = "HORA DE INGRESO";
            worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["E3"].Style.Font.Size = 10;
            worksheet.Cells["E3"].Style.WrapText = true;
            worksheet.Cells["E3"].Style.Font.Bold = true;
            worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["F3"].Value = "ESTADO";
            worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["F3"].Style.Font.Size = 10;
            worksheet.Cells["F3"].Style.WrapText = true;
            worksheet.Cells["F3"].Style.Font.Bold = true;
            worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["G3"].Value = "ASISTENCIA";
            worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["G3"].Style.Font.Size = 10;
            worksheet.Cells["G3"].Style.WrapText = true;
            worksheet.Cells["G3"].Style.Font.Bold = true;
            worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            int row = 4;

            foreach (DatosFormatoPersonaAsignacionExportModel rowitem in datos)
            {
                worksheet.Row(row).Height = 15.25;

                worksheet.Cells["A" + row].Value = rowitem.Persona;
                worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["A" + row].Style.Font.Size = 10;
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["B" + row].Value = rowitem.NombreCompleto;
                worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["B" + row].Style.Font.Size = 10;
                worksheet.Cells["B" + row].Style.WrapText = true;
                worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["C" + row].Value = rowitem.NombreArea;
                worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["C" + row].Style.Font.Size = 10;
                worksheet.Cells["C" + row].Style.WrapText = true;
                worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["D" + row].Value = rowitem.FechaAsignacion.ToString("dd/MM/yyyy");
                worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["D" + row].Style.Numberformat.Format = "yyyy-mm-dd";
                worksheet.Cells["D" + row].Style.Font.Size = 10;
                worksheet.Cells["D" + row].Style.WrapText = true;
                worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["E" + row].Value = rowitem.HoraIngreso.ToString("HH:mm");
                worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["E" + row].Style.Numberformat.Format = "HH:mm";
                worksheet.Cells["E" + row].Style.Font.Size = 10;
                worksheet.Cells["E" + row].Style.WrapText = true;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F" + row].Value = rowitem.Estado;
                worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["F" + row].Style.Font.Size = 10;
                worksheet.Cells["F" + row].Style.WrapText = true;
                worksheet.Cells["F" + row].Style.Font.Bold = true;
                worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G" + row].Value = rowitem.Asistencia;
                worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                worksheet.Cells["G" + row].Style.Font.Size = 10;
                worksheet.Cells["G" + row].Style.WrapText = true;
                worksheet.Cells["G" + row].Style.Font.Bold = true;
                worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (rowitem.Estado == "Activo")
                    worksheet.Cells["F" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#4b971c"));
                else
                    worksheet.Cells["F" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#8d0d43"));

                row++;
            }

        }

        private static void CuerpoListadoPersonal(ExcelPackage excelPackage, IEnumerable<DatosFormatoListadoPersonalAsignacion> datosListadoPersonal)
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Reporte listado Personal");
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.Cells.Style.Font.Size = 10;
            worksheet.Cells.Style.Font.Bold = true;

            worksheet.Cells["A1:P1"].Merge = true;
            worksheet.Cells["A1:P1"].Value = "INFORMACIÓN LISTADO DE PERSONAL";
            worksheet.Cells["A1:P1"].Style.Font.Size = 16;
            worksheet.Cells["A1:P1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ConfiguracionColumnaListadoPersonal(worksheet);

            worksheet.Cells["A3"].Value = "Fecha";
            worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A3"].Style.WrapText = true;
            worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["B3"].Value = "Clasificación Gerencia";
            worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B3"].Style.WrapText = true;
            worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["C3"].Value = "Cod. Persona";
            worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["C3"].Style.WrapText = true;
            worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["D3"].Value = "Nombre Completo";
            worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["D3"].Style.WrapText = true;
            worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["E3"].Value = "Horario";
            worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["E3"].Style.WrapText = true;
            worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["F3"].Value = "Area";
            worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["F3"].Style.WrapText = true;
            worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["G3"].Value = "Fecha de asignación";
            worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["G3"].Style.WrapText = true;
            worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["H3"].Value = "Fecha de reasignación";
            worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["H3"].Style.WrapText = true;
            worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["I3"].Value = "Estado";
            worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["I3"].Style.WrapText = true;
            worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["J3"].Value = "Ingreso";
            worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["J3"].Style.WrapText = true;
            worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["K3"].Value = "Salida";
            worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["K3"].Style.WrapText = true;
            worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["L3"].Value = "Hreal";
            worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["L3"].Style.WrapText = true;
            worksheet.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["M3"].Value = "Numero de dia";
            worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["M3"].Style.WrapText = true;
            worksheet.Cells["M3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["N3"].Value = "Tardanza";
            worksheet.Cells["N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N3"].Style.WrapText = true;
            worksheet.Cells["N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["O3"].Value = "Hextra";
            worksheet.Cells["O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["O3"].Style.WrapText = true;
            worksheet.Cells["O3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["P3"].Value = "Justificación";
            worksheet.Cells["P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["P3"].Style.WrapText = true;
            worksheet.Cells["P3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["Q3"].Value = "Vacaciones";
            worksheet.Cells["Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["Q3"].Style.WrapText = true;
            worksheet.Cells["Q3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            PintarCabeceraListadoPersonal(worksheet);
            int row = 4;
            foreach (DatosFormatoListadoPersonalAsignacion persona in datosListadoPersonal)
            {
                worksheet.Row(row).Height = 14.25;

                worksheet.Cells["A" + row].Value = persona.Fecha;
                worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B" + row].Value = persona.ClasificacionGerencia;
                worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + row].Style.WrapText = true;
                worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C" + row].Value = persona.Persona;
                worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + row].Style.WrapText = true;
                worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D" + row].Value = persona.NombreCompleto;
                worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + row].Style.WrapText = true;
                worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E" + row].Value = persona.Horario;
                worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + row].Style.WrapText = true;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F" + row].Value = persona.NombreArea;
                worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + row].Style.WrapText = true;
                worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G" + row].Value = persona.FechaAsignacion;
                worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["G" + row].Style.WrapText = true;
                worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["H" + row].Value = persona.FechaReAsignacion;
                worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["H" + row].Style.WrapText = true;
                worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["I" + row].Value = persona.Estado;
                worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + row].Style.WrapText = true;
                worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J" + row].Value = persona.Ingreso;
                worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + row].Style.WrapText = true;
                worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["K" + row].Value = persona.Salida;
                worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Style.WrapText = true;
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["L" + row].Value = persona.Hreal;
                worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["L" + row].Style.WrapText = true;
                worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["M" + row].Value = persona.NumDia;
                worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + row].Style.WrapText = true;
                worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["N" + row].Value = persona.Tardanza;
                worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N" + row].Style.WrapText = true;
                worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["O" + row].Value = persona.Hextra;
                worksheet.Cells["O" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O" + row].Style.WrapText = true;
                worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["P" + row].Value = persona.Justificaciones;
                worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P" + row].Style.WrapText = true;
                worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["Q" + row].Value = persona.Vacaciones;
                worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q" + row].Style.WrapText = true;
                worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row++;
            }
        }

        private static void ConfiguracionColumnaListadoPersonal(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 17.57 + 2.71;
            worksheet.Column(2).Width = 17.57 + 2.71;
            worksheet.Column(3).Width = 10 + 2.71;
            worksheet.Column(4).Width = 49.57 + 2.71;
            worksheet.Column(5).Width = 22.57 + 2.71;
            worksheet.Column(6).Width = 26.90 + 2.71;
            worksheet.Column(7).Width = 20.94 + 2.71;
            worksheet.Column(8).Width = 17.57 + 2.71;
            worksheet.Column(9).Width = 17.57 + 2.71;
            worksheet.Column(10).Width = 10.14 + 2.71;
            worksheet.Column(11).Width = 10.14 + 2.71;
            worksheet.Column(12).Width = 10.14 + 2.71;
            worksheet.Column(13).Width = 10.14 + 2.71;
            worksheet.Column(14).Width = 9.57 + 2.71;
            worksheet.Column(15).Width = 12.00 + 2.71;
            worksheet.Column(16).Width = 10.86 + 2.71;
            worksheet.Column(17).Width = 12.71 + 2.71;
            worksheet.Column(18).Width = 12.71 + 2.71;
        }


        private static void PintarCabeceraListadoPersonal(ExcelWorksheet worksheet)
        {
           worksheet.Cells["A3:Q3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#336ca5"));
           worksheet.Cells["A3:Q3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#fcffff"));
        }
    }
}
