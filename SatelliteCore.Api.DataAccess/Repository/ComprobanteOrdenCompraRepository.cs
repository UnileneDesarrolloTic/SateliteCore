using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Response.CompraAguja;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra;
using Dapper;
using SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ComprobanteOrdenCompraRepository : IComprobanteOrdenCompraRepository
    {
        private readonly IAppConfig _appConfig;

        public ComprobanteOrdenCompraRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }


        public async Task<IEnumerable<MostrarFechaPrometida>> MostrarInformacionOrdenCompra(string ordenCompra, string secuencia, string item)
        {
            IEnumerable<MostrarFechaPrometida> result = new List<MostrarFechaPrometida>();
            string sql = "SELECT NumeroDocumento, Secuencia, Item, Comentarios, Reprogramacion, Usuario FROM TBDHistorialFechaOrdenCompra WHERE Secuencia = @secuencia AND Item = @item AND NumeroDocumento = @ordenCompra;";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<MostrarFechaPrometida>(sql, new { ordenCompra, secuencia, item });
            }
            return result;
        }

        public async Task<string> RegistrarFechaPrometida(DatosFormatoRegistrarFecha dato, string usuario)
        {
   
            string sql = "INSERT INTO TBDHistorialFechaOrdenCompra (NumeroDocumento, Secuencia,  Item, Comentarios, Reprogramacion, Usuario) VALUES (@documento, @secuencia, @item, @comentario, @prometida, @usuario );";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(sql, new { dato.documento, dato.secuencia, dato.item, dato.comentario, dato.prometida, usuario });
            }
            return "";
        }
    }
}
