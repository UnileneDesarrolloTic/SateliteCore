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
using System.Linq;
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

            string sql = "INSERT INTO SatelliteCore.dbo.TBMDispensacionHistorica (EntregadoPor) VALUES (@usuario); SELECT SCOPE_IDENTITY();";
            int idDispensacion = 0;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                idDispensacion = await context.QueryFirstOrDefaultAsync<int>(sql, new { usuario });

                foreach (DatosFormatoDispensacionDetalleMP item in dato.detalleDispensacion)
                {
                    await context.ExecuteAsync("usp_satelite_registrar_dispensacion_MP", new { dato.ordenFabricacion, dato.itemTerminado ,item.secuencia, item.documento, item.itemInsumo, item.itemTipo, item.unidadCodigo, item.cantidadGeneral, item.cantidadSolicitada, item.cantidadDespachada, item.cantidadIngresada, item.tipoMP, item.lote, item.entregadoPor, item.recibidoPor, usuario, idDispensacion }, commandType: CommandType.StoredProcedure);
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

        public async Task<DatosFormatoInformacionDispensacionPT> DispensacionDetalleReceta(string item, string ordenFabricacion, string secuencia)
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

        public async Task<DatosFormatoDispensacionDetalle> DetalleDispensacionReceta()
        {
            DatosFormatoDispensacionDetalle result = new DatosFormatoDispensacionDetalle();
            using (SqlConnection connection = new SqlConnection(_appConfig.contextSpring))
            {
                using SqlMapper.GridReader multi = await connection.QueryMultipleAsync("usp_Satelite_DispensacionDetalle_Global", commandType: CommandType.StoredProcedure);
                result.DetalleDispensacion = multi.Read<DatosFormatoDispensacionRecetaDetalle>().ToList();
                result.SubFamilia = multi.Read<SubFamiliaDispensacion>().ToList();

            }
            return result;
        }

        public async Task<string> RegistrarRecetasGlobal(IEnumerable<DatosFormatoRegistroDispensacionRecetaGlobal> dato, string usuario)
        {
            string result = "";
            string sql = "INSERT INTO SatelliteCore.dbo.TBMDispensacionHistorica (EntregadoPor) VALUES (@usuario); SELECT SCOPE_IDENTITY();";
            int idDispensacion = 0;
            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                idDispensacion = await context.QueryFirstOrDefaultAsync<int>(sql, new { usuario }); 
                foreach (DatosFormatoRegistroDispensacionRecetaGlobal item in dato)
                {
                    await context.ExecuteAsync("usp_satelite_registrar_dispensacion_MP", new { item.ordenFabricacion, item.itemTerminado, item.secuencia, item.documento, item.itemInsumo, item.itemTipo, item.unidadCodigo, item.cantidadGeneral, item.cantidadSolicitada, item.cantidadDespachada, item.cantidadIngresada, item.tipoMP, item.lote, item.entregadoPor, item.recibidoPor, usuario, idDispensacion }, commandType: CommandType.StoredProcedure);
                }
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoDispensacionGuiaDespacho>> DispensacionGuiaDespacho(DatosFormatoFiltroDispensacion dato)
        {
            IEnumerable<DatosFormatoDispensacionGuiaDespacho> listado = new List<DatosFormatoDispensacionGuiaDespacho>();

            string sql = "SELECT Id, EntregadoPor, ISNULL(RecibidoPor, '') RecibidoPor, FechaRegistro, Estado, FechaDespacho FROM TBMDispensacionHistorica WHERE Estado = @estado " + (dato.fechaInicio == null ? "" : "AND CONVERT(varchar, FechaRegistro, 23)  BETWEEN @fechaInicio AND @fechaFin ");
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                listado = await context.QueryAsync<DatosFormatoDispensacionGuiaDespacho>(sql, new { dato.fechaInicio, dato.fechaFin, dato.estado });
                
            }
            return listado;
        }

        public async Task<IEnumerable<DatosFormatoMostrarDispensacionDespacho>> MostrarDispensacionDespacho(string id)
        {
            IEnumerable<DatosFormatoMostrarDispensacionDespacho> listado = new List<DatosFormatoMostrarDispensacionDespacho>();

            string sql = "SELECT IdDispensacion, OrdenFabricacion, ItemTerminado, Documento, Secuencia, a.Item, RTRIM(b.DescripcionLocal) DescripcionLocal, Lote, CantidadIngresada, UsuarioCreacion, FechaRegistro, CantidadSolicitada, IIF(a.Estado = 'PR', 'PREPARACION', 'APROBADO') Estado FROM TBDDispensacionHistorica a INNER JOIN PROD_UNILENE2..WH_ITEMMAST b ON  a.Item = b.Item WHERE Id = @id";
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                listado = await context.QueryAsync<DatosFormatoMostrarDispensacionDespacho>(sql, new { id});

            }
            return listado;
        }
    }
}
