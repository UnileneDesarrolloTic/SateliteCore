using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class LogisticaRepository : ILogisticaRepository
    {
        private readonly IAppConfig _appConfig;

        public LogisticaRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string NumeroGuia)
        {
            IEnumerable<DatosFormatoPlanOrdenServicosD> result = new List<DatosFormatoPlanOrdenServicosD>();

            string script = "SELECT CONCAT(RTRIM(a.SERIE),'-',RTRIM(a.NUMERO_DOCUMENTO))  NumeroGuia, FECHA_DOCUMENTO FechaDocumento, RTRIM(b.NombreCompleto) Cliente , " +
                            "RTRIM(a.FACTURA_NUMERO) OrdenServicios , a.FECHA_RETORNO FechaRetorno " +
                            "FROM UNILENE_REPORTEADOR..TLOG_PLAN_ORDEN_SERVICIO_D  a "+
                            "INNER JOIN PROD_UNILENE2..PersonaMast b ON a.CLIENTE = b.Persona "+
                            "WHERE a.ESTADO = 'A' AND a.NUMERO_DOCUMENTO = RIGHT('0000000000' + Ltrim(Rtrim(@NumeroGuia)), 10)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoPlanOrdenServicosD>(script, new { NumeroGuia });
            }

            return result;
        }


        public async Task<int> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato)
        {
            string script = "UPDATE UNILENE_REPORTEADOR..TLOG_PLAN_ORDEN_SERVICIO_D SET FECHA_RETORNO=@fechaRetorno  WHERE " +
                            "CONCAT(RTRIM(SERIE),'-',RTRIM(NUMERO_DOCUMENTO))=@numeroGuia";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script,dato);
            }

            return 1;
        }
    }
}
