using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Encajado;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class EncajadoRespository: IEncajadoRespository
    {
        private readonly IAppConfig _appConfig;

        public EncajadoRespository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }


        public async Task<List<ListaOrdenesFabricaciónDTO>> ListaOrdenesFabricacion(string ordenFabricacion)
        {
            IEnumerable<ListaOrdenesFabricaciónDTO> lista = new List<ListaOrdenesFabricaciónDTO>();

            string query = "SELECT TOP 20 a.NumeroLote OrdenFabricacion, a.Item, b.DescripcionLocal Descripcion, a.CantidadProgramada, a.CantidadProducida, " +
                "a.CantidadMuestra, a.Estado, a.FechaExpiracion " +
                "FROM EP_ProgramacionLote a WITH(NOLOCK) INNER JOIN WH_ItemMast b WITH(NOLOCK) ON a.ITEM = b.Item " +
                "WHERE a.CompaniaSocio = '01000000'";

            if (!string.IsNullOrWhiteSpace(ordenFabricacion))
                query = query + " AND a.NUMEROLOTE = @ordenFabricacion";

            query = query + " ORDER BY FechaProduccion DESC";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await context.QueryAsync<ListaOrdenesFabricaciónDTO>(query, new { ordenFabricacion });
            }

            return lista.ToList();
        }
    }
}
