using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;
using SatelliteCore.Api.Models.Response.Logistica;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class DispensacionRepository : IDispensacionRepository
    {
        private readonly IAppConfig _appConfig;

        public DispensacionRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }



        public async Task<IEnumerable<DatosFormatoObtenerOrdenFabricacion>> ObtenerOrdenFabricacion(DatosFormatoFiltroOrdenFabricacion dato)
        {
            IEnumerable<DatosFormatoObtenerOrdenFabricacion> result = new List<DatosFormatoObtenerOrdenFabricacion>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatoObtenerOrdenFabricacion>("usp_listaOrdenFabricacion", new { dato.fechaInicio, dato.fechaFinal, dato.lote, dato.ordenFabricacion }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion>> RecetasOrdenFabricacion(string ordenFabricacion)
        {
            IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion> result = new List<DatosFormatoListadoMateriaPrimaDispensacion>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatoListadoMateriaPrimaDispensacion>("usp_listaOrdenFabricacionReceta", new { ordenFabricacion }, commandType: CommandType.StoredProcedure);
            }
            return result;
        }

        public async Task<string> RegistrarDispensacionMP(DatosFormatoDispensacionMateriaPrima dato, string usuario)
        {
            string result = "";
            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {    foreach (DatosFormatoDispensacionDetalleMP item in dato.detalleDispensacion)
                {
                    await context.ExecuteAsync("usp_satelite_registrar_dispensacion_MP", new { dato.ordenFabricacion, dato.itemTerminado ,item.secuencia, item.documento, item.itemInsumo, item.itemTipo, item.unidadCodigo, item.cantidadGeneral, item.cantidadSolicitada, item.cantidadDespachada, item.cantidadIngresada, item.tipoMP, item.lote, item.entregadoPor, item.recibidoPor, usuario }, commandType: CommandType.StoredProcedure);
                }
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoHistorialDispensaccion>> HistorialDispensacionMP(string ordenFabricacion, string lote)
        {
            
            IEnumerable<DatosFormatoHistorialDispensaccion> listado = new List<DatosFormatoHistorialDispensaccion>();
            string sql = "SELECT Documento, Secuencia, a.Item, RTRIM(b.DescripcionLocal) Descripcion,  Lote, CantidadSolicitada, CantidadIngresada, UsuarioCreacion, FechaRegistro FROM TBDDispensacionHistorica a INNER JOIN PROD_UNILENE2..WH_ItemMast b ON a.Item = b.Item WHERE OrdenFabricacion = IIF(@ordenFabricacion = '', OrdenFabricacion, @ordenFabricacion) " +
                         " OR Lote = IIF(@lote = '', Lote, @lote) ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                listado = await context.QueryAsync<DatosFormatoHistorialDispensaccion>(sql, new { ordenFabricacion, lote });

            }
            return listado;
        }

        public async Task<DatosFormatoInformacionDispensacionPT> InformacionItem(string item, string ordenFabricacion, string secuencia)
        {
            DatosFormatoInformacionDispensacionPT resultItem = new DatosFormatoInformacionDispensacionPT();

            string query = "SELECT a.NUMEROLOTE OrdenFabricacion, RTRIM( a.Item) Item, RTRIM(b.descripcionLocal) Descripcion, (a.CANTIDADPROGRAMADA +  ISNULL(a.CANTIDADMUESTRA,0)) CantidadTotal, dp.Cantidad CantidadParcial " +
                           "FROM EP_PROGRAMACIONLOTE a " +
                           "INNER JOIN WH_ITEMMAST b ON a.ITEM = b.ITEM " +
                           "INNER join SatelliteCore..TBDDividirProgramacion dp ON a.NUMEROLOTE = dp.ordenfabricacion and a.REFERENCIANUMERO = dp.lote and dp.Estado = 'A' " +
                           "WHERE a.ITEM = @item AND a.NUMEROLOTE = @ordenFabricacion AND dp.Secuencia = @secuencia AND dp.Estado = 'A'";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                resultItem = await connection.QueryFirstOrDefaultAsync<DatosFormatoInformacionDispensacionPT>(query, new { item, ordenFabricacion, secuencia });
                connection.Dispose();
            }

            return resultItem;
        }

    }
}
