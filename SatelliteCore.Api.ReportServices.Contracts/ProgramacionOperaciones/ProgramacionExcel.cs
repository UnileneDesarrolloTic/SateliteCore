using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.ProgramacionOperaciones;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace SatelliteCore.Api.ReportServices.Contracts.ProgramacionOperaciones
{
    public class ProgramacionExcel
    {
        public string Programacion(IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion> listaProgramado, IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion> listaNoProgramado)
        {

            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                Exportacion(excelPackage, listaProgramado, "Programado");
                Exportacion(excelPackage, listaNoProgramado,  "No Programado");

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

                return reporte;

            }
        }
        private static void Exportacion(ExcelPackage excelPackage, IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion> listaProgramado, string hojaTitulo)
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(hojaTitulo);
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.Cells.Style.Font.Size = 9;
            worksheet.View.ZoomScale = 100;
            
            int fila = 4;

            ConfiguracionTamanioCeldasProgramacion(worksheet);
            PintarCabeceraProgramacion(worksheet, fila, hojaTitulo);

            

            worksheet.Cells["A" + fila].Value = "Orden Fabricacion";
            worksheet.Cells["A" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A" + fila].Style.WrapText = true;
            worksheet.Cells["A" + fila].Style.Font.Bold = true;
            worksheet.Cells["A" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["B" + fila].Value = "Lote";
            worksheet.Cells["B" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B" + fila].Style.WrapText = true;
            worksheet.Cells["B" + fila].Style.Font.Bold = true;
            worksheet.Cells["B" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["C" + fila].Value = "Fecha Producción";
            worksheet.Cells["C" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["C" + fila].Style.WrapText = true;
            worksheet.Cells["C" + fila].Style.Font.Bold = true;
            worksheet.Cells["C" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["D" + fila].Value = "Item";
            worksheet.Cells["D" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["D" + fila].Style.WrapText = true;
            worksheet.Cells["D" + fila].Style.Font.Bold = true;
            worksheet.Cells["D" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["D" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["E" + fila].Value = "Descripción";
            worksheet.Cells["E" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["E" + fila].Style.WrapText = true;
            worksheet.Cells["E" + fila].Style.Font.Bold = true;
            worksheet.Cells["E" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["E" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["F" + fila].Value = "Cantidad Programación";
            worksheet.Cells["F" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["F" + fila].Style.WrapText = true;
            worksheet.Cells["F" + fila].Style.Font.Bold = true;
            worksheet.Cells["F" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["F" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["G" + fila].Value = "Cantidad Pedida";
            worksheet.Cells["G" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["G" + fila].Style.WrapText = true;
            worksheet.Cells["G" + fila].Style.Font.Bold = true;
            worksheet.Cells["G" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["G" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["H" + fila].Value = "Tipo";
            worksheet.Cells["H" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["H" + fila].Style.WrapText = true;
            worksheet.Cells["H" + fila].Style.Font.Bold = true;
            worksheet.Cells["H" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["H" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["I" + fila].Value = "Fecha Requerida";
            worksheet.Cells["I" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["I" + fila].Style.WrapText = true;
            worksheet.Cells["I" + fila].Style.Font.Bold = true;
            worksheet.Cells["I" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["I" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["J" + fila].Value = "Fecha Inicio";
            worksheet.Cells["J" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["J" + fila].Style.WrapText = true;
            worksheet.Cells["J" + fila].Style.Font.Bold = true;
            worksheet.Cells["J" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["J" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["K" + fila].Value = "Fecha Final";
            worksheet.Cells["K" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["K" + fila].Style.WrapText = true;
            worksheet.Cells["K" + fila].Style.Font.Bold = true;
            worksheet.Cells["K" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["K" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["L" + fila].Value = "Cliente";
            worksheet.Cells["L" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["L" + fila].Style.WrapText = true;
            worksheet.Cells["L" + fila].Style.Font.Bold = true;
            worksheet.Cells["L" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["L" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            int row = fila + 1;
            foreach (DatosFormatoProgramacionOperacionesOrdenFabricacion item in listaProgramado)
            {
                worksheet.Row(row).Height = 14.25;

                worksheet.Cells["A" + row].Value = item.OrdenFabricacion;
                worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B" + row].Value = item.Lote;
                worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + row].Style.WrapText = true;
                worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C" + row].Value = item.FechaProduccion;
                worksheet.Cells["C" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + row].Style.WrapText = true;
                worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D" + row].Value = item.Item;
                worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + row].Style.WrapText = true;
                worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E" + row].Value = item.DescripcionLocal;
                worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + row].Style.WrapText = true;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F" + row].Value = item.CantidadProgramada;
                worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + row].Style.WrapText = true;
                worksheet.Cells["F" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                worksheet.Cells["G" + row].Value = item.CantidadPedida;
                worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + row].Style.WrapText = true;
                worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["H" + row].Value = item.ReferenciaTipo;
                worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + row].Style.WrapText = true;
                worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["I" + row].Value = item.FechaRequerida;
                worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + row].Style.WrapText = true;
                worksheet.Cells["I" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J" + row].Value = item.FechaInicio;
                worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + row].Style.WrapText = true;
                worksheet.Cells["J" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["K" + row].Value = item.FechaEntrega;
                worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Style.WrapText = true;
                worksheet.Cells["K" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["L" + row].Value = item.Busqueda;
                worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L" + row].Style.WrapText = true;
                worksheet.Cells["L" + row].Style.Numberformat.Format = "dd/MM/YYYY";
                worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                row++;
            }
        }

        private static void PintarCabeceraProgramacion(ExcelWorksheet worksheet, int fila,string hojaTitulo)
        {   
            if(hojaTitulo == "No Programado")
            {
                worksheet.Cells["A" + fila + ":L" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#5589c4"));
                worksheet.Cells["A" + fila + ":L" + fila].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            }
            else
            {
                worksheet.Cells["A" + fila + ":L" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#dc143c"));
                worksheet.Cells["A" + fila + ":L" + fila].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            }
          
        }


        private static void ConfiguracionTamanioCeldasProgramacion(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 12.86 + 2.71;
            worksheet.Column(2).Width = 11.71 + 2.71;
            worksheet.Column(3).Width = 11.71 + 2.71;
            worksheet.Column(4).Width = 17.14 + 2.71;
            worksheet.Column(5).Width = 65.00 + 2.71;
            worksheet.Column(6).Width = 12.57 + 2.71;
            worksheet.Column(7).Width = 10.86 + 2.71;
            worksheet.Column(8).Width = 10 + 2.71;
            worksheet.Column(9).Width = 15.57 + 2.71;
            worksheet.Column(10).Width = 15.57 + 2.71;
            worksheet.Column(11).Width = 15.86 + 2.71;
            worksheet.Column(12).Width = 35.00 + 2.71;
        }
    }
}
