using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class LicitacionesRepository : ILicitacionesRepository
    {
        private readonly IAppConfig _appConfig;
        public LicitacionesRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido)
        {

            IEnumerable<ListarDetallePedido> result = new List<ListarDetallePedido>();
            string script = "SELECT RTRIM(a.NumeroDocumento) Pedido, a.Linea, RTRIM(a.ItemCodigo) Item, RTRIM(a.Descripcion) Descripcion, a.CantidadPedida , a.PrecioUnitario, a.Monto " +
                            "FROM CO_DocumentoDetalle a INNER JOIN CO_Documento b ON  a.CompaniaSocio=b.CompaniaSocio AND a.TipoDocumento=b.TipoDocumento AND a.NumeroDocumento=b.NumeroDocumento " +
                            "WHERE  a.TipoDocumento = 'PE' AND  a.NumeroDocumento = @Pedido  AND b.ClienteNumero=2317";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<ListarDetallePedido>(script, new { Pedido });
            }
            return result;
        }

        public async Task<int> RegistrarProceso(DatoFormatoProcesoModel dato)
        {
            int result;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql1 = "INSERT INTO TBMLIProceso(DescripcionProceso, Cliente)" +
                             "VALUES(@Proceso, @Cliente)";

                result  = await context.ExecuteAsync(sql1, new { dato.Proceso, dato.Cliente });
                        //await context.ExecuteAsync(sql2, analisis);

            }
            return result;
        }

    }
}
