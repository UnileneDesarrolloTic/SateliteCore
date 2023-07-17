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
                result = await context.QueryAsync<DatosFormatoObtenerOrdenFabricacion>("usp_listaOrdenFabricacion", new { dato.fechaInicio, dato.fechaFinal, dato.lote, dato.ordenFabricacion, dato.estado }, commandType: CommandType.StoredProcedure);
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

    }
}
