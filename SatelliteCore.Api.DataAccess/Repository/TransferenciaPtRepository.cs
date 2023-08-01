using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Response.TransferenciaPT;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class TransferenciaPtRepository : ITransferenciaPtRepository
    {
        private readonly IAppConfig _appConfig;

        public TransferenciaPtRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<List<PendienteTransFisicaDTO>> ListaPendienteTransfereciaFisica(string almacenCodigo)
        {
            IEnumerable<PendienteTransFisicaDTO> result = new List<PendienteTransFisicaDTO>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<PendienteTransFisicaDTO>("SP_ListarPendientesTransferenciaFisica", new { almacenCodigo }, commandType: CommandType.StoredProcedure);
            }

            return result.ToList();
        }

        public async Task RegistraTransfenciaPT(int idControl, string controlNumero, decimal cantidadTotal, decimal cantidadParcial, string usuario)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB  ))
            {
                await context.QueryAsync<PendienteTransFisicaDTO>("SP_InsertarTransferenciaFisica", 
                        new {idControl, controlNumero, cantidadTotal, cantidadParcial, usuarioTraslado = usuario}, commandType: CommandType.StoredProcedure);
            }
        }

        public async Task<List<PendienteRecepcionarPtDTO>> ListaPendienteRecepcionFisica(string almacenCodigo, string estado)
        {
            IEnumerable<PendienteRecepcionarPtDTO> result = new List<PendienteRecepcionarPtDTO>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<PendienteRecepcionarPtDTO>("SP_ListaRecepcionarTransferenciaPT", new { almacenCodigo, estado }, commandType: CommandType.StoredProcedure);
            }

            return result.ToList();
        }

        public async Task<string> RegistraRecepcionPT (RegistrarRecepcionPtDTO recepcion)
        {
            string mensaje = "";
            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                mensaje  = await context.QueryFirstOrDefaultAsync<string>("SP_RecepcionarStockMP", recepcion, commandType: CommandType.StoredProcedure);
            }

            return mensaje;
        }

    }
}
