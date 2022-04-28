using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;
namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    class Formato70_Report
    {

        public static string Exportar(Image firma,Image logoUnilene, Coti_Formato_70_Model cotizacion)
        {
            byte[] file;

            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Cotización Essalud Juliaca F-2");

                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("Unilene", logoUnilene);
                imagenUnilene.SetPosition(0, 0, 0, 4);
                imagenUnilene.SetSize(200, 50);

                workSheet.Cells.Style.Font.Name = "Arial";
                workSheet.Cells.Style.Font.Size = 8;
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);


                ConfigurarTamanioDeCeldas(workSheet);
                UnirCeldas(workSheet);
                PintarCeldas(workSheet);
                BordesCeldas(workSheet);



                workSheet.Cells["A4"].Value = "RED ASISTENCIAL JULIACA - ESSALUD";
                workSheet.Cells["A4"].Style.Font.Size = 8;
                workSheet.Cells["A4"].Style.Font.Name = "Arial";
                workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A4"].Style.Font.Bold = true;

                workSheet.Cells["A5"].Value = "DIVISIÓN DE ADQUISICIONES";
                workSheet.Cells["A5"].Style.Font.Size = 8;
                workSheet.Cells["A5"].Style.Font.Name = "Arial";
                workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A5"].Style.Font.Bold = true;


                workSheet.Cells["A7"].Value = "DECLARACION JURADA DE COTIZACION: ADQUISICION DE MATERIAL MEDICO DELEGADO 1er TRIMESTRE 2022 - ESSALUD";
                workSheet.Cells["A7"].Style.Font.Size = 8;
                workSheet.Cells["A7"].Style.Font.Name = "Arial";
                workSheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A7"].Style.Font.Bold = true;


                workSheet.Cells["A9"].Value = "FECHA DE COTIZACION: " + cotizacion.Prov_Fecha.ToLongDateString() ;
                workSheet.Cells["A9"].Style.Font.Size = 14;
                workSheet.Cells["A9"].Style.Font.Name = "Arial";
                workSheet.Cells["A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A9"].Style.Font.Bold = true;

                workSheet.Cells["A11"].Value = "RAZON SOCIAL ";
                workSheet.Cells["A11"].Style.Font.Size = 8;
                workSheet.Cells["A11"].Style.Font.Name = "Arial";
                workSheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                //workSheet.Cells["A11"].Style.Border.BorderAround(Exc);

                workSheet.Cells["C11"].Value = cotizacion.Prov_RazonSocial;
                workSheet.Cells["C11"].Style.Font.Size = 8;
                workSheet.Cells["C11"].Style.Font.Name = "Arial";
                workSheet.Cells["C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["O11"].Value = "REPRESENTANTE ";
                workSheet.Cells["O11"].Style.Font.Size = 8;
                workSheet.Cells["O11"].Style.Font.Name = "Arial";
                workSheet.Cells["O11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
               

                workSheet.Cells["Q11"].Value = cotizacion.Prov_Representante;
                workSheet.Cells["Q11"].Style.Font.Size = 8;
                workSheet.Cells["Q11"].Style.Font.Name = "Arial";
                workSheet.Cells["Q11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["A12"].Value = "RUC ";
                workSheet.Cells["A12"].Style.Font.Size = 8;
                workSheet.Cells["A12"].Style.Font.Name = "Arial";
                workSheet.Cells["A12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["C12"].Value = cotizacion.Prov_Ruc;
                workSheet.Cells["C12"].Style.Font.Size = 8;
                workSheet.Cells["C12"].Style.Font.Name = "Arial";
                workSheet.Cells["C12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["O12"].Value = "CARGO ";
                workSheet.Cells["O12"].Style.Font.Size = 8;
                workSheet.Cells["O12"].Style.Font.Name = "Arial";
                workSheet.Cells["O12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["Q12"].Value = cotizacion.Prov_Cargo;
                workSheet.Cells["Q12"].Style.Font.Size = 8;
                workSheet.Cells["Q12"].Style.Font.Name = "Arial";
                workSheet.Cells["Q12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["A13"].Value = "DIRECCION ";
                workSheet.Cells["A13"].Style.Font.Size = 8;
                workSheet.Cells["A13"].Style.Font.Name = "Arial";
                workSheet.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["C13"].Value = cotizacion.Prov_Direccion;
                workSheet.Cells["C13"].Style.Font.Size = 8;
                workSheet.Cells["C13"].Style.Font.Name = "Arial";
                workSheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["O13"].Value = "EMAIL ";
                workSheet.Cells["O13"].Style.Font.Size = 8;
                workSheet.Cells["O13"].Style.Font.Name = "Arial";
                workSheet.Cells["O13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["Q13"].Value = cotizacion.Prov_Email2;
                workSheet.Cells["Q13"].Style.Font.Size = 8;
                workSheet.Cells["Q13"].Style.Font.Name = "Arial";
                workSheet.Cells["Q13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["A14"].Value = "TELEFONO (S) ";
                workSheet.Cells["A14"].Style.Font.Size = 8;
                workSheet.Cells["A14"].Style.Font.Name = "Arial";
                workSheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["C14"].Value = cotizacion.Prov_Telefono;
                workSheet.Cells["C14"].Style.Font.Size = 8;
                workSheet.Cells["C14"].Style.Font.Name = "Arial";
                workSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["O14"].Value = "CELULAR ";
                workSheet.Cells["O14"].Style.Font.Size = 8;
                workSheet.Cells["O14"].Style.Font.Name = "Arial";
                workSheet.Cells["O14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["Q14"].Value = cotizacion.Prov_movil;
                workSheet.Cells["Q14"].Style.Font.Size = 8;
                workSheet.Cells["Q14"].Style.Font.Name = "Arial";
                workSheet.Cells["Q14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;



                workSheet.Cells["A15"].Value = "E-MAIL";
                workSheet.Cells["A15"].Style.Font.Size = 8;
                workSheet.Cells["A15"].Style.Font.Name = "Arial";
                workSheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["C15"].Value = cotizacion.Prov_Email1;
                workSheet.Cells["C15"].Style.Font.Size = 8;
                workSheet.Cells["C15"].Style.Font.Name = "Arial";
                workSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["O15"].Value = "RPM / RPC";
                workSheet.Cells["O15"].Style.Font.Size = 8;
                workSheet.Cells["O15"].Style.Font.Name = "Arial";
                workSheet.Cells["O15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["Q15"].Value = cotizacion.Prov_Rpm;
                workSheet.Cells["O15"].Style.Font.Size = 8;
                workSheet.Cells["O15"].Style.Font.Name = "Arial";
                workSheet.Cells["O15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["A16"].Value = "VALIDEZ  DE LA OFERTA";
                workSheet.Cells["A16"].Style.Font.Size = 8;
                workSheet.Cells["A16"].Style.Font.Name = "Arial";
                workSheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["C16"].Value = cotizacion.Prov_valiOferta;
                workSheet.Cells["C16"].Style.Font.Size = 8;
                workSheet.Cells["C16"].Style.Font.Name = "Arial";
                workSheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["O16"].Value = "CONDICION DE PAGO";
                workSheet.Cells["O16"].Style.Font.Size = 7;
                workSheet.Cells["O16"].Style.Font.Name = "Arial";
                workSheet.Cells["O16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                workSheet.Cells["Q16"].Value = cotizacion.Prov_CondPago;
                workSheet.Cells["Q16"].Style.Font.Size = 8;
                workSheet.Cells["Q16"].Style.Font.Name = "Arial";
                workSheet.Cells["Q16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                workSheet.Cells["A18"].Value = "SIRVASE LLENAR LA SOLICITUD DE COTIZACION PARA COMPRA DE: ";
                workSheet.Cells["A18"].Style.Font.Size = 8;
                workSheet.Cells["A18"].Style.Font.Name = "Arial";
                workSheet.Cells["A18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A18"].Style.Font.Bold = true;

                workSheet.Cells["A20"].Value = "N° ITEM";
                workSheet.Cells["A20"].Style.Font.Size = 10;
                workSheet.Cells["A20"].Style.Font.Name = "Arial";
                workSheet.Cells["A20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["A20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A20"].Style.Font.Bold = true;
                workSheet.Cells["A20"].Style.WrapText = true;


                workSheet.Cells["B20"].Value = "REQUERMIENTO";
                workSheet.Cells["B20"].Style.Font.Size = 10;
                workSheet.Cells["B20"].Style.Font.Name = "Arial";
                workSheet.Cells["B20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B20"].Style.Font.Bold = true;
                workSheet.Cells["B20"].Style.WrapText = true;

                workSheet.Cells["F20"].Value = "COTIZACION";
                workSheet.Cells["F20"].Style.Font.Size = 10;
                workSheet.Cells["F20"].Style.Font.Name = "Arial";
                workSheet.Cells["F20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F20"].Style.Font.Bold = true;
                workSheet.Cells["F20"].Style.WrapText = true;

                workSheet.Cells["M20"].Value = "REQUISITOS TÉCNICOS MÍNIMOS Y CONDICIONES GEN. P/ADQUISICIÓN DE DISPOSITIVO MEDICO EN ESSALUD";
                workSheet.Cells["M20"].Style.Font.Size = 10;
                workSheet.Cells["M20"].Style.Font.Name = "Arial";
                workSheet.Cells["M20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M20"].Style.Font.Bold = true;
                workSheet.Cells["M20"].Style.WrapText = true;

                workSheet.Cells["B21"].Value = "CODIGO SAP";
                workSheet.Cells["B21"].Style.Font.Size = 10;
                workSheet.Cells["B21"].Style.Font.Name = "Arial";
                workSheet.Cells["B21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["B21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["B21"].Style.Font.Bold = true;
                workSheet.Cells["B21"].Style.WrapText = true;

                workSheet.Cells["C21"].Value = "DENOMINACION";
                workSheet.Cells["C21"].Style.Font.Size = 10;
                workSheet.Cells["C21"].Style.Font.Name = "Arial";
                workSheet.Cells["C21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["C21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["C21"].Style.Font.Bold = true;
                workSheet.Cells["C21"].Style.WrapText = true;

                workSheet.Cells["D21"].Value = "UM";
                workSheet.Cells["D21"].Style.Font.Size = 10;
                workSheet.Cells["D21"].Style.Font.Name = "Arial";
                workSheet.Cells["D21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["D21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["D21"].Style.Font.Bold = true;
                workSheet.Cells["D21"].Style.WrapText = true;

                workSheet.Cells["E21"].Value = "CANTIDAD";
                workSheet.Cells["E21"].Style.Font.Size = 10;
                workSheet.Cells["E21"].Style.Font.Name = "Arial";
                workSheet.Cells["E21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["E21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["E21"].Style.Font.Bold = true;
                workSheet.Cells["E21"].Style.WrapText = true;

                workSheet.Cells["F21"].Value = "PRECIO UNITARIO S/.";
                workSheet.Cells["F21"].Style.Font.Size = 10;
                workSheet.Cells["F21"].Style.Font.Name = "Arial";
                workSheet.Cells["F21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["F21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["F21"].Style.Font.Bold = true;
                workSheet.Cells["F21"].Style.WrapText = true;
                workSheet.Cells["F21"].Style.TextRotation = 90;


                workSheet.Cells["G21"].Value = "VALOR TOTAL S/.";
                workSheet.Cells["G21"].Style.Font.Size = 10;
                workSheet.Cells["G21"].Style.Font.Name = "Arial";
                workSheet.Cells["G21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["G21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["G21"].Style.Font.Bold = true;
                workSheet.Cells["G21"].Style.WrapText = true;
                workSheet.Cells["G21"].Style.TextRotation = 90;

                workSheet.Cells["H21"].Value = "MARCA";
                workSheet.Cells["H21"].Style.Font.Size = 10;
                workSheet.Cells["H21"].Style.Font.Name = "Arial";
                workSheet.Cells["H21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["H21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["H21"].Style.Font.Bold = true;
                workSheet.Cells["H21"].Style.WrapText = true;
                workSheet.Cells["H21"].Style.TextRotation = 90;

                workSheet.Cells["I21"].Value = "PROCEDENCIA";
                workSheet.Cells["I21"].Style.Font.Size = 10;
                workSheet.Cells["I21"].Style.Font.Name = "Arial";
                workSheet.Cells["I21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["I21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["I21"].Style.Font.Bold = true;
                workSheet.Cells["I21"].Style.WrapText = true;
                workSheet.Cells["I21"].Style.TextRotation = 90;

                workSheet.Cells["J21"].Value = "FORMA DE PRESENTACION";
                workSheet.Cells["J21"].Style.Font.Size = 10;
                workSheet.Cells["J21"].Style.Font.Name = "Arial";
                workSheet.Cells["J21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["J21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["J21"].Style.Font.Bold = true;
                workSheet.Cells["J21"].Style.WrapText = true;
                workSheet.Cells["J21"].Style.TextRotation = 90;

                workSheet.Cells["K21"].Value = "VIGENCIA DEL PRODUCTO";
                workSheet.Cells["K21"].Style.Font.Size = 10;
                workSheet.Cells["K21"].Style.Font.Name = "Arial";
                workSheet.Cells["K21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["K21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["K21"].Style.Font.Bold = true;
                workSheet.Cells["K21"].Style.WrapText = true;
                workSheet.Cells["K21"].Style.TextRotation = 90;

                workSheet.Cells["L21"].Value = "PLAZO DE ENTREGA";
                workSheet.Cells["L21"].Style.Font.Size = 10;
                workSheet.Cells["L21"].Style.Font.Name = "Arial";
                workSheet.Cells["L21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["L21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["L21"].Style.Font.Bold = true;
                workSheet.Cells["L21"].Style.WrapText = true;
                workSheet.Cells["L21"].Style.TextRotation = 90;

                workSheet.Cells["M21"].Value = "Registro Sanitario y/o Certificado de Registro Sanitario vigente (Si-No)";
                workSheet.Cells["M21"].Style.Font.Size = 10;
                workSheet.Cells["M21"].Style.Font.Name = "Arial";
                workSheet.Cells["M21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["M21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["M21"].Style.Font.Bold = true;
                workSheet.Cells["M21"].Style.WrapText = true;
                workSheet.Cells["M21"].Style.TextRotation = 90;

                workSheet.Cells["N21"].Value = "Protocolo y/o Ceretiifcado de Análisis de Producto Terminado (Si o No)";
                workSheet.Cells["N21"].Style.Font.Size = 10;
                workSheet.Cells["N21"].Style.Font.Name = "Arial";
                workSheet.Cells["N21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N21"].Style.Font.Bold = true;
                workSheet.Cells["N21"].Style.WrapText = true;
                workSheet.Cells["N21"].Style.TextRotation = 90;

                workSheet.Cells["O21"].Value = "Metodología de Análisis Propia ó a que farmacopea se acoge (Si-No)";
                workSheet.Cells["O21"].Style.Font.Size = 10;
                workSheet.Cells["O21"].Style.Font.Name = "Arial";
                workSheet.Cells["O21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["O21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["O21"].Style.Font.Bold = true;
                workSheet.Cells["O21"].Style.WrapText = true;
                workSheet.Cells["O21"].Style.TextRotation = 90;

                workSheet.Cells["P21"].Value = "Resolucion de Autorización Sanitaria o Constancia de Registro de Establecimiento Farmaceutico (Si o No)";
                workSheet.Cells["P21"].Style.Font.Size = 10;
                workSheet.Cells["P21"].Style.Font.Name = "Arial";
                workSheet.Cells["P21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["P21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["P21"].Style.Font.Bold = true;
                workSheet.Cells["P21"].Style.WrapText = true;
                workSheet.Cells["P21"].Style.TextRotation = 90;


                workSheet.Cells["Q21"].Value = "Certificado de Buenas Prácticas de Manufactura (CBPM) Vigente (Si o No)";
                workSheet.Cells["Q21"].Style.Font.Size = 10;
                workSheet.Cells["Q21"].Style.Font.Name = "Arial";
                workSheet.Cells["Q21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["Q21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["Q21"].Style.Font.Bold = true;
                workSheet.Cells["Q21"].Style.WrapText = true;
                workSheet.Cells["Q21"].Style.TextRotation = 90;


                workSheet.Cells["R21"].Value = "Certificado de Buena Práctica de Almacenamiento (CBPA) vigente (Si - No) ";
                workSheet.Cells["R21"].Style.Font.Size = 10;
                workSheet.Cells["R21"].Style.Font.Name = "Arial";
                workSheet.Cells["R21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["R21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["R21"].Style.Font.Bold = true;
                workSheet.Cells["R21"].Style.WrapText = true;
                workSheet.Cells["R21"].Style.TextRotation = 90;


                workSheet.Cells["S21"].Value = "Cumple al 100% con las Caracteristicas Técnicos denominación: Descripción del Item  (Si-No)";
                workSheet.Cells["S21"].Style.Font.Size = 10;
                workSheet.Cells["S21"].Style.Font.Name = "Arial";
                workSheet.Cells["S21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["S21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["S21"].Style.Font.Bold = true;
                workSheet.Cells["S21"].Style.WrapText = true;
                workSheet.Cells["S21"].Style.TextRotation = 90;


                int index = 22;

                foreach (Coti_Formato70_Detalle item in cotizacion.Detalle)
                {
                    workSheet.Row(index).Height = 25;
                    workSheet.Cells["A" + index].Value = index - 21;
                    workSheet.Cells["A" + index].Style.Font.Size = 10;
                    workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["A" + index].Style.Numberformat.Format = "0";
                    workSheet.Cells["A" + index].Style.WrapText = true;

                    workSheet.Cells["B" + index].Value = item.CodigoSAP;
                    workSheet.Cells["B" + index].Style.Font.Size = 10;
                    workSheet.Cells["B" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["B" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["B" + index].Style.WrapText = true;


                    workSheet.Cells["C" + index].Value = item.Denominacion;
                    workSheet.Cells["C" + index].Style.Font.Size = 10;
                    workSheet.Cells["C" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["C" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["C" + index].Style.WrapText = true;


                    workSheet.Cells["D" + index].Value = item.Um;
                    workSheet.Cells["D" + index].Style.Font.Size = 10;
                    workSheet.Cells["D" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["D" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["D" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["D" + index].Style.WrapText = true;


                    workSheet.Cells["E" + index].Value = item.Cantidad;
                    workSheet.Cells["E" + index].Style.Font.Size = 10;
                    workSheet.Cells["E" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["E" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["E" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["E" + index].Style.WrapText = true;
                    workSheet.Cells["E" + index].Style.Numberformat.Format = "#,##0";

                    workSheet.Cells["F" + index].Value = item.PreUnitario;
                    workSheet.Cells["F" + index].Style.Font.Size = 10;
                    workSheet.Cells["F" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["F" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["F" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["F" + index].Style.WrapText = true;
                    workSheet.Cells["F" + index].Style.Numberformat.Format = "#,##0.00";

                    workSheet.Cells["G" + index].Value = item.PreTotal;
                    workSheet.Cells["G" + index].Style.Font.Size = 10;
                    workSheet.Cells["G" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["G" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["G" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["G" + index].Style.WrapText = true;
                    workSheet.Cells["G" + index].Style.Numberformat.Format = "#,##0.00";


                    workSheet.Cells["H" + index].Value = item.Marca;
                    workSheet.Cells["H" + index].Style.Font.Size = 10;
                    workSheet.Cells["H" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["H" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["H" + index].Style.WrapText = true;


                    workSheet.Cells["I" + index].Value = item.Procedencia;
                    workSheet.Cells["I" + index].Style.Font.Size = 10;
                    workSheet.Cells["I" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["I" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["I" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["I" + index].Style.WrapText = true;

                    workSheet.Cells["J" + index].Value = item.Presentacion;
                    workSheet.Cells["J" + index].Style.Font.Size = 10;
                    workSheet.Cells["J" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["J" + index].Style.WrapText = true;

                    workSheet.Cells["K" + index].Value = item.VigProducto;
                    workSheet.Cells["K" + index].Style.Font.Size = 10;
                    workSheet.Cells["K" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["K" + index].Style.WrapText = true;


                    workSheet.Cells["L" + index].Value = item.PlazoEntrega;
                    workSheet.Cells["L" + index].Style.Font.Size = 10;
                    workSheet.Cells["L" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["L" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["L" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["L" + index].Style.WrapText = true;

                    workSheet.Cells["M" + index].Value = item.RegSanitario;
                    workSheet.Cells["M" + index].Style.Font.Size = 10;
                    workSheet.Cells["M" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["M" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["M" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["M" + index].Style.WrapText = true;


                    workSheet.Cells["N" + index].Value = item.CProtocoloTerminado;
                    workSheet.Cells["N" + index].Style.Font.Size = 10;
                    workSheet.Cells["N" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["N" + index].Style.WrapText = true;

                    workSheet.Cells["O" + index].Value = item.Metodoanalisis;
                    workSheet.Cells["O" + index].Style.Font.Size = 10;
                    workSheet.Cells["O" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["O" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["O" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["O" + index].Style.WrapText = true;


                    workSheet.Cells["P" + index].Value = item.RAutoSanitaria;
                    workSheet.Cells["P" + index].Style.Font.Size = 10;
                    workSheet.Cells["P" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["P" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["P" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["P" + index].Style.WrapText = true;


                    workSheet.Cells["Q" + index].Value = item.Cbpm;
                    workSheet.Cells["Q" + index].Style.Font.Size = 10;
                    workSheet.Cells["Q" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["Q" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["Q" + index].Style.WrapText = true;

                    workSheet.Cells["R" + index].Value = item.Cbpa;
                    workSheet.Cells["R" + index].Style.Font.Size = 10;
                    workSheet.Cells["R" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["R" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["R" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["R" + index].Style.WrapText = true;


                    workSheet.Cells["S" + index].Value = item.Ccditem;
                    workSheet.Cells["S" + index].Style.Font.Size = 10;
                    workSheet.Cells["S" + index].Style.Font.Name = "Arial";
                    workSheet.Cells["S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["S" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    workSheet.Cells["S" + index].Style.WrapText = true;



                    index++;
                }

                index++;
                workSheet.Cells["A" + index + ":C" + index ].Merge =true;
                workSheet.Cells["A" + index].Value = "CONDICIONES GENERALES";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Bold = true;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Cells["A" + index + ":S"+ index].Merge = true;
                workSheet.Cells["A" + index].Value = "1.- El precio de Mercado será a todo costo, es decir, deberá incluir todos los tributos (incluido el I.G.V.), seguros, transportes, inspecciones";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 42;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "2.- Cabe precisar que frente a un reclamo, queja u observación de los materiales medicos, queda en potestad de Essalud, el sometimiento al control de calidad respectivo, cuyo costo será asumido enteramente por el proveedor, en caso el resultado sea no conforme el plazo para el canje de los saldos inmovilizados en cada alamacen de la entidad es de 15 dias calendarios contados a partir de recibida la comunicación escrita por parte de la entidad.";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;


                index++;
                workSheet.Row(index).Height = 30.75;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "3.- De haberse efectuado el pago de un lote declarado no conforme, el contratista se obliga a reponer el costo total de las cantidades consumidas y en caso de no efectuarse el canje abonara el costo total del lote inmovilizado a LA ENTIDAD mediante pago en efectivo o cheque de gerencia o deduciendolo de cualqueira de sus facturas";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Entregas:";
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Mensual";
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":S" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C5D9F1"));


                index++;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Lugar de Entrega:";
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "La entregas se realizarán en el almacen de la Redes Asistencial Juliaca, sito en Av. Jose Santos Chocano S/N - Urb. La Capilla";
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":S" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C5D9F1"));

                index++;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "Plazo de Entrega:";
                workSheet.Cells["A" + index].Style.Font.UnderLine = true;
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;

                index++;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["A" + index + ":S" + index].Merge = true;
                workSheet.Cells["A" + index].Value = "El plazo de entrega es días calendarios según cuadro precedente después de recepcionada la Orden de Compra. Si el plazo propuesto es superior a 10 días, estará sujeto a evaluación por EsSalud.";
                
                workSheet.Cells["A" + index].Style.Font.Size = 10;
                workSheet.Cells["A" + index].Style.Font.Name = "Arial";
                workSheet.Cells["A" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                workSheet.Cells["A" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells["A" + index].Style.WrapText = true;
                workSheet.Cells["A" + index + ":S" + index].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C5D9F1"));


                index = index + 5;
                workSheet.Row(index).Height = 20.25;
                workSheet.Cells["N" + index + ":R" + index].Merge = true;
                workSheet.Cells["N" + index].Value = "Fima y Sello del Representante Legal";
                workSheet.Cells["N" + index].Style.Font.Size = 12;
                workSheet.Cells["N" + index].Style.Font.Name = "Arial";
                workSheet.Cells["N" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells["N" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells["N" + index].Style.WrapText = true;

                ExcelPicture firmaCatizacion = workSheet.Drawings.AddPicture("Firma_Unilene", firma);
                firmaCatizacion.SetPosition(index-6, -8, 13, 85);
                firmaCatizacion.SetSize(230,100);



                workSheet.View.ZoomScale = 80; //80
                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

            }
            return reporte;
        }


        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet workSheet)
        {

            workSheet.Column(1).Width = 6.00 + 0.71;
            workSheet.Column(2).Width = 13.71 + 0.71;
            workSheet.Column(3).Width = 48.43 + 0.71;
            workSheet.Column(4).Width = 6.00 + 0.71;
            workSheet.Column(5).Width = 10.00 + 0.71;
            workSheet.Column(6).Width = 10 + 0.71;
            workSheet.Column(7).Width = 10 + 0.71;
            workSheet.Column(8).Width = 13.43 + 0.71;
            workSheet.Column(9).Width = 9.57 + 0.71;
            workSheet.Column(10).Width = 9.57 + 0.71;
            workSheet.Column(11).Width = 9.57 + 0.71;
            workSheet.Column(12).Width = 7.29 + 0.71;
            workSheet.Column(13).Width = 9.29 + 0.71;
            workSheet.Column(14).Width = 9.29 + 0.71;
            workSheet.Column(15).Width = 8.14 + 0.71;
            workSheet.Column(16).Width = 8.14 + 0.71;
            workSheet.Column(17).Width = 7.29 + 0.71;
            workSheet.Column(18).Width = 7.29 + 0.71;
            workSheet.Column(19).Width = 7.29 + 0.71;



            workSheet.Row(1).Height = 12.75;
            workSheet.Row(2).Height = 12.75;
            workSheet.Row(3).Height = 12.75;
            workSheet.Row(4).Height = 12.75;
            workSheet.Row(5).Height = 12.75;
            workSheet.Row(6).Height = 12.75;

            workSheet.Row(7).Height = 42.00;
            workSheet.Row(8).Height = 12.75;
            workSheet.Row(9).Height = 18.00;

            workSheet.Row(10).Height = 13.50;
            workSheet.Row(11).Height = 13.50;
            workSheet.Row(12).Height = 13.50;
            workSheet.Row(13).Height = 13.50;
            workSheet.Row(14).Height = 13.50;
            workSheet.Row(15).Height = 13.5;
            workSheet.Row(16).Height = 13.5;
            workSheet.Row(17).Height = 6.75;
            workSheet.Row(18).Height = 13.5;
            workSheet.Row(19).Height = 5.25;
            workSheet.Row(20).Height = 24.00;
            workSheet.Row(21).Height = 246.75;
            
        }
        private static void UnirCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A4:C4"].Merge = true;
            workSheet.Cells["A5:C5"].Merge = true;
            workSheet.Cells["A7:S7"].Merge = true;
            workSheet.Cells["A9:S9"].Merge = true;

            workSheet.Cells["A11:B11"].Merge = true;
            workSheet.Cells["C11:H11"].Merge = true;
            workSheet.Cells["O11:P11"].Merge = true;
            workSheet.Cells["Q11:S11"].Merge = true;


            workSheet.Cells["A12:B12"].Merge = true;
            workSheet.Cells["C12:H12"].Merge = true;
            workSheet.Cells["O12:P12"].Merge = true;
            workSheet.Cells["Q12:S12"].Merge = true;

            workSheet.Cells["A13:B13"].Merge = true;
            workSheet.Cells["C13:H13"].Merge = true;
            workSheet.Cells["O12:P12"].Merge = true;
            workSheet.Cells["Q12:S12"].Merge = true;

            workSheet.Cells["A14:B14"].Merge = true;
            workSheet.Cells["C14:H14"].Merge = true;
            workSheet.Cells["O14:P14"].Merge = true;
            workSheet.Cells["Q14:S14"].Merge = true;

            workSheet.Cells["A15:B15"].Merge = true;
            workSheet.Cells["C15:H15"].Merge = true;
            workSheet.Cells["O15:P15"].Merge = true;
            workSheet.Cells["Q15:S15"].Merge = true;


            workSheet.Cells["A16:B16"].Merge = true;
            workSheet.Cells["C16:H16"].Merge = true;
            workSheet.Cells["O16:P16"].Merge = true;
            workSheet.Cells["Q16:S16"].Merge = true;


            workSheet.Cells["A18:S18"].Merge = true;


            workSheet.Cells["A20:A21"].Merge = true;
            workSheet.Cells["B20:E20"].Merge = true;
            workSheet.Cells["F20:L20"].Merge = true;
            workSheet.Cells["M20:S20"].Merge = true;
        }

        private static void BordesCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A7:S7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["A20:A21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B20:E20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F20:L20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M20:S20"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["B21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["C21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["D21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["E21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["F21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["G21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["H21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["I21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["J21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["K21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["L21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["M21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["N21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["O21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["P21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["Q21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["R21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            workSheet.Cells["S21"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
           
        }
        private static void PintarCeldas(ExcelWorksheet workSheet)
        {
            workSheet.Cells["A7:S7"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C0C0C0"));
        }
    }
}
