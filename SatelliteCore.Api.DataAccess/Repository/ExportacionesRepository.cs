using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ExportacionesRepository : IExportacionesRepository
    {
        private readonly IAppConfig _appConfig;
        public ExportacionesRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<DatosFormatoListarCotizacionExportacion>> ListarCotizacionExportaciones(FiltrarCotizacionExportacionModel filtro)
        {
            IEnumerable<DatosFormatoListarCotizacionExportacion> result = new List<DatosFormatoListarCotizacionExportacion>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoListarCotizacionExportacion>("usp_listar_Cotizacion_Exportaciones", new { filtro.NumeroDocumento, filtro.Estado, filtro.Cliente, filtro.FechaInicio, filtro.FechaFin }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<(object cabecera, object detalle)> BuscarCotizacionExportaciones(string NumeroDocumento)
        {
            (object cabecera, object detalle) datosCotizacion;

            string script = " SELECT RTRIM(NumeroDocumento) NumeroDocumento, ClienteNumero, RTRIM(ClienteNombre) ClienteNombre,RTRIM(ClienteRUC) ClienteRUC,ISNULL(RTRIM(ClienteDireccion),'') ClienteDireccion,RTRIM(LugarEntrega) LugarEntrega,RTRIM(Comentarios) Comentarios, Contacto ,  MontoAfecto,MontoNoAfecto , MontoDescuentos , MontoImpuestoVentas,  MontoRedondeo , MontoTotal FROM PROD_UNILENE2..CO_Cotizacion WHERE NumeroDocumento = @NumeroDocumento " +
                            " SELECT RTRIM(b.NumeroDeParte) codSut,RTRIM(a.ItemCodigo) item, RTRIM(a.Descripcion) descripcion,a.CantidadPedida cantidad, a.PrecioUnitario precioUnitario, a.Monto  monto, IIF(a.IGVExoneradoFlag = 'N', CAST(0 AS BIT), CAST(1 AS BIT)) iGVExoneradoFlag  " +
                            " FROM PROD_UNILENE2..CO_CotizacionDetalle a " +
                            " INNER JOIN PROD_UNILENE2..WH_ItemMast b ON a.ItemCodigo = b.Item " +
                            " WHERE NumeroDocumento =  @NumeroDocumento";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using var result = await context.QueryMultipleAsync(script, new { NumeroDocumento });
                datosCotizacion.cabecera = result.Read().FirstOrDefault();
                datosCotizacion.detalle = result.Read().ToList();
            }

            return datosCotizacion;
        }

        public async Task<int> EditarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones dato, string UsuarioSesion)
        {
            int result = 1;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync("usp_Editar_Cotizacion_Exportacion_Cabecera", new { dato.Contacto, dato.Comentarios, dato.NumeroDocumento, dato.MontoAfecto, dato.MontoNoAfecto , dato.Descuento , dato.ImpVentas , dato.AjusteRedondeo , dato.MontoTotal, UsuarioSesion }, commandType: CommandType.StoredProcedure);

                foreach (FormatoDetalleCotizacionExportaciones item in dato.DetalleCotizacion)
                {
                    await connection.ExecuteAsync("usp_Editar_Cotizacion_Exportacion_Detalle", new { item.cantidad, item.precioUnitario, item.monto, item.iGVExoneradoFlag, item.item, dato.NumeroDocumento }, commandType: CommandType.StoredProcedure);
                }

                connection.Dispose();
            }

            return result;
        }

        public async Task<int> RegistrarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones dato, string UsuarioSesion)
        {
            int response = 0;
            string NumeroDocumento = "";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                NumeroDocumento = await connection.QueryFirstOrDefaultAsync<string>("usp_Insertar_Cotizacion_Header", new { dato.Cliente, dato.MontoAfecto, dato.MontoNoAfecto, dato.Descuento, dato.ImpVentas,dato.MontoTotal,dato.AjusteRedondeo,dato.Comentarios, UsuarioSesion }, commandType: CommandType.StoredProcedure);
                int Linea = 0;
                foreach (FormatoDetalleCotizacionExportaciones item in dato.DetalleCotizacion)  
                {   
                    await connection.ExecuteAsync("usp_Insertar_Cotizacion_Detalle", new { NumeroDocumento, Linea , item.item, item.cantidad, item.precioUnitario, item.monto, item.iGVExoneradoFlag, UsuarioSesion }, commandType: CommandType.StoredProcedure);
                    Linea ++;
                }

                connection.Dispose();
            }

            return response;
        }

        public async Task<FormatoDetalleCotizacionExportaciones> ObtenerInformacionExcel(FormatoDetalleExcelExportacionesModel parametro)
        {
            FormatoDetalleCotizacionExportaciones result_db = new FormatoDetalleCotizacionExportaciones();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result_db = await connection.QueryFirstOrDefaultAsync<FormatoDetalleCotizacionExportaciones>("usp_Buscar_Items_Cotizacion_Exportaciones", new { parametro.Codsut, parametro.Cantidad, parametro.Punitario }, commandType: CommandType.StoredProcedure);

                connection.Dispose();
            }

            return result_db;
        }
    }
}
