using OfficeOpenXml;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.Logistica;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class ExportacionesServices : IExportacionesServices
    {
        private readonly IExportacionesRepository _exportacionesRepository;

        public ExportacionesServices(IExportacionesRepository exportacionesRepository)
        {
            _exportacionesRepository = exportacionesRepository;
        }

        public async Task<IEnumerable<DatosFormatoListarCotizacionExportacion>> ListarCotizacionExportaciones(FiltrarCotizacionExportacionModel filtro)
        {
            return await _exportacionesRepository.ListarCotizacionExportaciones(filtro);
        }
        public async Task<(object cabecera, object detalle)> BuscarCotizacionExportaciones(string NumeroDocumento)
        {
            (object cabecera, object detalle) response = await _exportacionesRepository.BuscarCotizacionExportaciones(NumeroDocumento);
            return response;
        }
        public async Task<ResponseModel<string>> GuardarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones datos, string UsuarioSesion)
        {
            if (datos.FormularioNuevo == false)
                  await _exportacionesRepository.RegistrarCotizacionExportaciones(datos, UsuarioSesion);
            else
                  await _exportacionesRepository.EditarCotizacionExportaciones(datos, UsuarioSesion);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, datos.FormularioNuevo ==false ? "Registrado con exito" : "Modificación con exito" );
        }

        public async Task<ResponseModel<List<FormatoDetalleCotizacionExportaciones>>> ProcesarExcelExportaciones(DatosFormato64 dato)
        {

            byte[] byteArray = Convert.FromBase64String(dato.base64string);

            ResponseModel<List<FormatoDetalleCotizacionExportaciones>> datosArchivos;

            using (MemoryStream memStream = new MemoryStream(byteArray))
            {
                using (ExcelPackage package = new ExcelPackage(memStream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var sheet = package.Workbook.Worksheets.First();
                    datosArchivos = await GetList(sheet);
                }

            }

            return datosArchivos;
        }

        private async Task<ResponseModel<List<FormatoDetalleCotizacionExportaciones>>> GetList(ExcelWorksheet sheet)
        {
            List<FormatoDetalleCotizacionExportaciones> list = new List<FormatoDetalleCotizacionExportaciones>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() }
            );
            var longuitud = columnInfo.Count();

            if (longuitud != 3)
                return new ResponseModel<List<FormatoDetalleCotizacionExportaciones>>(false, "El Formato de Columna es codsut,cantidad,unitario", null);


            var startRow = sheet.Dimension.Start.Row;
            var endRow = sheet.Dimension.End.Row;

            for (int row = 2; row <= sheet.Dimension.Rows; row++)
            {
                try
                {
                    FormatoDetalleExcelExportacionesModel parametro = new FormatoDetalleExcelExportacionesModel();

                    if(sheet.Cells[row, 1].Value.ToString().Length!=21)
                        return new ResponseModel<List<FormatoDetalleCotizacionExportaciones>>(false, "Revisar la fila de codsut ( " + sheet.Cells[row, 1].Value.ToString()  + ") tienes " + sheet.Cells[row, 1].Value.ToString().Length +" caracteres", null);

                    parametro.Codsut = sheet.Cells[row, 1].Value.ToString();
                    parametro.Cantidad = int.Parse(sheet.Cells[row, 2].Value.ToString());
                    parametro.Punitario = Convert.ToDecimal(sheet.Cells[row, 3].Value.ToString());

                    FormatoDetalleCotizacionExportaciones obtenerinformacion = new FormatoDetalleCotizacionExportaciones();

                    obtenerinformacion = await _exportacionesRepository.ObtenerInformacionExcel(parametro);

                    list.Add(obtenerinformacion);
                }
                catch (Exception)
                {

                    return new ResponseModel<List<FormatoDetalleCotizacionExportaciones>>(false, "Formato incorrecto, Revisar Excel", null);
                }

                   
            }
            return new ResponseModel<List<FormatoDetalleCotizacionExportaciones>>(true, Constante.MESSAGE_SUCCESS, list);
        }

        public async Task<ResponseModel<List<FormatoDetalleCotizacionExportaciones>>> BuscarWHItemMast(string  Opcion, string Descripcion)
        {

            List <FormatoDetalleCotizacionExportaciones> result = new List<FormatoDetalleCotizacionExportaciones>();
            result = await _exportacionesRepository.BuscarWHItemMast(Opcion,Descripcion);

            return new ResponseModel<List<FormatoDetalleCotizacionExportaciones>>(true, Constante.MESSAGE_SUCCESS, result);
        }

        public async Task<ResponseModel<string>> DesactivarItemCotizacionExportacion(string NumeroDocumento, string Item, int Linea, string UsuarioSesion)
        {
             await _exportacionesRepository.DesactivarItemCotizacionExportacion(NumeroDocumento, Item, Linea ,UsuarioSesion);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Eliminación exitosa ");
        }


    }
}
