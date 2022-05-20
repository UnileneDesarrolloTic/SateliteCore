using Dapper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ContabilidadController : ControllerBase
    {
        private readonly IContabilidadService _ContabilidadService;

        public ContabilidadController(IContabilidadService ContabilidadService)
        {
            _ContabilidadService = ContabilidadService;

        }

        [HttpGet("ListarDetraccionContabilidad")]
        public async Task<ActionResult> ListarDetraccion()
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, "Los datos enviado no son válidos.", "");

                return BadRequest(responseError);
            }

            List<DetraccionesEntity> response = await _ContabilidadService.ListarDetraccion();
            return Ok(response);
        }

        [HttpGet("ProcesarDetraccionContabilidad")]
        public async Task<ActionResult> ProcesarDetraccionContabilidad(string urlarchivo)
        {
            string file = @"C:\"+ urlarchivo;
            List<FormatoComprobantePagoDetraccion> datosArchivos;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["UNILENE"];
                datosArchivos = GetList<FormatoComprobantePagoDetraccion>(sheet);
            }

            int response = await _ContabilidadService.ProcesarDetraccionContabilidad(datosArchivos);
 
            return Ok(response);
        }

        private List<FormatoComprobantePagoDetraccion> GetList<T>(ExcelWorksheet sheet)
        {
            List<FormatoComprobantePagoDetraccion> list = new List<FormatoComprobantePagoDetraccion>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() }
            );

            var startRow = sheet.Dimension.Start.Row;
            var endRow = sheet.Dimension.End.Row;

            for (int row = 2; row <= sheet.Dimension.Rows; row++)
            {
                FormatoComprobantePagoDetraccion parameter = new FormatoComprobantePagoDetraccion();
                parameter.TipoCuenta = sheet.Cells[row, 1].Value.ToString();
                parameter.NumeroCuenta = sheet.Cells[row, 2].Value.ToString();
                parameter.NumeroConstancia = sheet.Cells[row, 3].Value.ToString();
                parameter.PeriodoTributario = sheet.Cells[row, 4].Value.ToString();
                parameter.RucProveedor = sheet.Cells[row, 5].Value.ToString();
                parameter.NombreProveedor = sheet.Cells[row, 6].Value.ToString();
                parameter.TipoDocumento = sheet.Cells[row, 7].Value.ToString();
                parameter.DocumentoAdquiriente = sheet.Cells[row, 8].Value.ToString();
                parameter.RazonSocial = sheet.Cells[row, 9].Value.ToString();
                parameter.FechaPago = new DateTime(Convert.ToDateTime(sheet.Cells[row, 10].Value).Year, Convert.ToDateTime(sheet.Cells[row, 10].Value).Month, Convert.ToDateTime(sheet.Cells[row, 10].Value).Day);
                parameter.MontoDeposito = Convert.ToDecimal(sheet.Cells[row, 11].Value.ToString());
                parameter.TipoBien = sheet.Cells[row, 12].Value.ToString();
                parameter.TipoOperacion = sheet.Cells[row, 13].Value.ToString();
                parameter.TipodeComprobante = sheet.Cells[row, 14].Value.ToString();
                parameter.Serie = sheet.Cells[row, 15].Value.ToString();
                parameter.Numero = sheet.Cells[row, 16].Value.ToString();
                parameter.PagoDetraccion = sheet.Cells[row, 17].Value.ToString();

                list.Add(parameter);
            }
            return list;
        }

        [HttpPost("GernerarBlogNotasDetraccion")]
        public string GenerarBlogNotasDetraccion(FormatoProcesoDetracciones dato)
        {
            string valortext = "P20197705249UNILENE SAC                        " + dato.periodo + dato.totalimporte.ToString("D15") + "\n";

            foreach (var detraccion in dato.proceso)
            {
                var numero = Int32.Parse(detraccion.Numero);
                var servicio = Int32.Parse(detraccion.BienServicio);
                var cuenta = "00002007770";
                var operacion = Int32.Parse(detraccion.Tipo);
                var tipo = Int32.Parse(detraccion.Tipo);

                valortext += detraccion.TipoDocumento + "" + detraccion.Ruc + "                                   " + servicio.ToString("D12") + cuenta + detraccion.Importe.ToString("D13") + operacion.ToString("D4") + detraccion.Periodo + tipo.ToString("D2") + detraccion.Serie + numero.ToString("D8") + "\n";
            }
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(valortext);

            return System.Convert.ToBase64String(plainTextBytes);
        }


    }
}