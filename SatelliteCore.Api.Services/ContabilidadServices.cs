
using OfficeOpenXml;
using System;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Services
{
    public class ContabilidadServices : IContabilidadService
    {
        private readonly IContabilidadRepository _contabilidadRepository;

        public ContabilidadServices(IContabilidadRepository contabilidadRepository)
        {
            _contabilidadRepository = contabilidadRepository;
        }
        public async Task<IEnumerable<DetraccionesEntity>> ListarDetraccion()
        {

            IEnumerable<DetraccionesEntity> lista  = await _contabilidadRepository.ListarDetraccion();
            return lista;
        }

        public int  ProcesarDetraccionContabilidad (string urlarchivo)
        {

            string file = @"C:\" + urlarchivo;
            List<FormatoComprobantePagoDetraccion> datosArchivos;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["UNILENE"];
                datosArchivos = GetList<FormatoComprobantePagoDetraccion>(sheet);
            }

            int response = _contabilidadRepository.ProcesarDetraccionContabilidad(datosArchivos);
            return response;
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





    }
    }
