
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
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.ReportServices.Contracts.Detracciones;
using System.Text;

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

        public int ProcesarDetraccionContabilidad (DatosFormato64 dato)
        {

            int response = 0;

            byte[] byteArray = Convert.FromBase64String(dato.base64string);

            List<FormatoComprobantePagoDetraccion> datosArchivos;

            using (MemoryStream memStream = new MemoryStream(byteArray))
            {
                using (ExcelPackage package = new ExcelPackage(memStream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var sheet = package.Workbook.Worksheets["UNILENE"];
                    datosArchivos = GetList<FormatoComprobantePagoDetraccion>(sheet);
                }
            }

            response = _contabilidadRepository.ProcesarDetraccionContabilidad(datosArchivos);

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
                parameter.PagoDetraccion = sheet.Cells[row, 17].Value == null ? "" : sheet.Cells[row, 17].Value.ToString();

                list.Add(parameter);
            }
            return list;
        }

        public string GenerarBlogNotasDetraccion(FormatoProcesoDetracciones dato)
        {
            GenerarBlogNotas Detraccion = new GenerarBlogNotas();
            string reporte =  Detraccion.ProcesarGenerarBlogNotas(dato);

            return reporte;
        }

        public async Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ConsultarProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato)
        {
            IEnumerable<DatosFormatoDatosProductoCostobase> lista = await _contabilidadRepository.ConsultarProductoCostoBase(dato);
            return lista;
        }



        public async Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ProcesarProductoExcel(DatosFormatoFiltrarAnalisisCostoRequest dato)
        {

            byte[] byteArray = Convert.FromBase64String(dato.base64);

            string ListarItem;

            using (MemoryStream memStream = new MemoryStream(byteArray))
            {
                using (ExcelPackage package = new ExcelPackage(memStream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var sheet = package.Workbook.Worksheets["UNILENE"];
                    ListarItem = GetListItem<string>(sheet);
                }
            }

            dato.base64 = ListarItem;

            IEnumerable<DatosFormatoDatosProductoCostobase>  Listar = await _contabilidadRepository.ConsultarProductoCostoBase(dato);

            return Listar;
        }
        private string GetListItem<T>(ExcelWorksheet sheet)
        {
            List<string> list = new List<string>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() }
            );

            var startRow = sheet.Dimension.Start.Row;
            var endRow = sheet.Dimension.End.Row;
            var Respuesta = "";

            StringBuilder builder = new StringBuilder();

            for (int row = 2; row <= endRow; row++)
            {
                builder.Append(sheet.Cells[row, 1].Value).Append(",");
            }
            Respuesta = builder.ToString();

            return Respuesta;
        }


        public async Task<IEnumerable<DatosFormatoRecetaItemComponente>> ConsultarRecetaProducto(string Item)
        {
            IEnumerable<DatosFormatoRecetaItemComponente> lista = await _contabilidadRepository.ConsultarRecetaProducto(Item);
            return lista;
        }

        public async Task<IEnumerable<DatosFormatoComponentePrecioUnitario>> ListarItemComponentePrecio(DatosFormatosComponentPrecio dato)
        {
            IEnumerable<DatosFormatoComponentePrecioUnitario> lista = await _contabilidadRepository.ListarItemComponentePrecio(dato);
            return lista;
        }




    }
}
