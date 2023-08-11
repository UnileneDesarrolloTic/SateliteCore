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
   
            string sql = "INSERT INTO TBDHistorialFechaOrdenCompra (NumeroDocumento, Secuencia,  Item, Comentarios, Reprogramacion, Usuario) VALUES (@ordenCompra, @secuencia, @item, @comentario, @prometida, @usuario );";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {   
                foreach(DatosFormatoDetalleOrdenCompraRequest item in dato.detalle)
                await context.ExecuteAsync(sql, new { dato.ordenCompra, item.secuencia, item.item, dato.comentario, dato.prometida, usuario });
            }
            return "";
        }


        public async Task<IEnumerable<DatosFormatoDetalleOrdenCompra>> MostrarDetalleOrdenCompra(string ordenCompra, string item, string secuencia)
        {
            IEnumerable<DatosFormatoDetalleOrdenCompra> listado = new List<DatosFormatoDetalleOrdenCompra>();

            string sql = ";WITH Temp_OrdenCompraItem AS (" +
                         " SELECT CAST(1 AS BIT) Seleccionar, NumeroOrden Documento, RTRIM(Item) Item, Secuencia, RTRIM(Descripcion) Descripcion " +
                         " FROM WH_OrdenCompraDetalle WHERE NumeroOrden = @ordenCompra AND Secuencia = @secuencia AND  Item = @item) " +
                         " SELECT ISNULL(b.Seleccionar, CAST(0 AS BIT)) Seleccionar, a.NumeroOrden Documento, RTRIM(a.Item) Item, a.Secuencia, RTRIM(a.Descripcion) Descripcion  " +
                         " FROM WH_OrdenCompraDetalle a LEFT JOIN Temp_OrdenCompraItem b ON a.NumeroOrden = b.Documento AND a.Item = b.Item AND a.Secuencia = b.Secuencia " +
                         " WHERE NumeroOrden = @ordenCompra  AND Estado IN ('PE','PR') ORDER BY Secuencia ASC ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
               listado = await context.QueryAsync<DatosFormatoDetalleOrdenCompra>(sql, new { ordenCompra, item, secuencia });
            }
            return listado;
        }
    }
}
