using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ProduccionRepository : IProduccionRepository
    {
        private readonly IAppConfig _appConfig;

        public ProduccionRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<SeguimientoProductoArimaModel> SeguimientoProductosArima(string periodo)
        {
            SeguimientoProductoArimaModel result = new SeguimientoProductoArimaModel();

            using (SqlConnection springContext = new SqlConnection(_appConfig.contextSpring))
            {
                using (SqlMapper.GridReader multi = await springContext.QueryMultipleAsync("usp_Satelite_ProductoTerminadoArima", new { periodo }, commandType: CommandType.StoredProcedure))
                {
                    result.Productos = multi.Read<ProductoArimaModel>().ToList();
                    result.DetalleTransito = multi.Read<TransitoProductoArimaModel>().ToList();
                }
            }

            return result;
        }

        public async Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla)
        {
            SeguimientoCandMPAGenericModel result = new SeguimientoCandMPAGenericModel();

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                result.SeguimientoCandidatosMPA = await satelliteContext.QueryAsync<SeguimientoCandMPAModel>("usp_pro_SeguimientoCandidatoMPA", new { regla }, commandType: CommandType.StoredProcedure);
                result.OrdenComprasPendientes = await satelliteContext.QueryAsync<DetalleSeguimientoCandMPAModel>("usp_pro_SeguimientoDetalleCandidatoMPA", new { regla }, commandType: CommandType.StoredProcedure);

                satelliteContext.Dispose();
            }

            return result;
        }

        public async Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro)
        {

            (IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros) result;
            DynamicParameters parametros = new DynamicParameters();

            parametros.Add("@FechaInicio", filtro.FechaInicio);
            parametros.Add("@FechaFin", filtro.FechaFin);
            parametros.Add("@Item", filtro.Item);
            parametros.Add("@Pagina", filtro.Pagina);
            parametros.Add("@RegistrosPorPagina", filtro.RegistrosPorPagina);
            parametros.Add("@TotalRegistos", dbType: DbType.Int32, direction: ParameterDirection.Output);

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                result.ListaPedidos = await satelliteContext.QueryAsync<PedidosCreadosAutoLogModel>("usp_Pro_ListarPedidosCreadosAuto", parametros, commandType: CommandType.StoredProcedure);
                result.TotalRegistros = parametros.Get<int>("@TotalRegistos");
                satelliteContext.Dispose();
            }

            return result;
        }

        public async Task<SeguimientoComprasMPArima> SeguimientoCompraMPArima(PronosticoCompraMP dato)
        {
            SeguimientoComprasMPArima result = new SeguimientoComprasMPArima();

            using (SqlConnection DMVentasContext = new SqlConnection(_appConfig.contextSpring))
            {
                using (SqlMapper.GridReader multi = await DMVentasContext.QueryMultipleAsync("usp_Satelite_CompraMateriaPrimaArima", dato , commandType: CommandType.StoredProcedure))
                {
                    result.Productos = multi.Read<CompraMPArimaModel>().ToList();
                    result.DetalleTransito = multi.Read<DCompraMPArimaModel>().ToList();
                    result.DetalleCalidad = multi.Read<CompraMPArimaDetalleControlCalidad>().ToList();

                }

            }

            return result;
        }
    }
}
