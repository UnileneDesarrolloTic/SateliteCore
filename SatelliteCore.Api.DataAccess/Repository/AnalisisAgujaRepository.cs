using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
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
    public class AnalisisAgujaRepository : IAnalisisAgujaRepository
    {
        private readonly IAppConfig _appConfig;

        public AnalisisAgujaRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ListarAnalisisAgujaModel>> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina)
        {
            IEnumerable<ListarAnalisisAgujaModel> result = new List<ListarAnalisisAgujaModel>();

            string script = "SELECT Lote, Item, DescripcionItem, OrdenCompra, Proveedor, Cantidad, FechaRegistro, CantidadPruebas FROM TBMAnalisisAgujas AA " +
                "WHERE 1=1" + (string.IsNullOrEmpty(ordenCompra) ? "" : " AND OrdenCompra LIKE @ordenCompra") + (string.IsNullOrEmpty(lote) ? "" : " AND Lote LIKE @lote") +
                (string.IsNullOrEmpty(item) ? "" : " AND DescripcionItem LIKE @item") + " ORDER BY FechaRegistro DESC OFFSET (@pagina - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<ListarAnalisisAgujaModel>(script, new { ordenCompra, lote, item, pagina });
            }

            return result;
        }

        public async Task<IEnumerable<ListarOrdenCompra>> ListaOrdenesCompra(string NumeroOrden)
        {
            IEnumerable<ListarOrdenCompra> result = new List<ListarOrdenCompra>();
            string script = "SELECT a.Linea Secuencia,RTRIM(b.ControlNumero) ControlNumero,RTRIM(ISNULL(c.Item, a.Item)) Item,c.Descripcion DescripcionItem,RTRIM(c.UnidadCodigo) UnidadCodigo," +
                "CAST(c.CantidadPedida AS INT) CantidadPedida,CAST(c.CantidadRecibida AS INT) CantidadRecibida,ISNULL(e.Lote, aa.LoteAprobado) AS LoteAprobado," +
                "ISNULL(aa.LoteRechazado, '') AS LoteRechazado,d.Proveedor CodProveedor,RTRIM(b.NumeroOrden) NumeroOrden, ISNULL(e.Lote, '') Analisis FROM WH_ControlCalidadDetalle a WITH(NOLOCK) " +
                "INNER JOIN WH_ControlCalidad b WITH(NOLOCK) ON a.CompaniaSocio = b.CompaniaSocio AND a.ControlNumero = b.ControlNumero " +
                "LEFT JOIN WH_ControlCalidadDetalle aa WITH(NOLOCK) ON aa.CompaniaSocio = a.CompaniaSocio AND aa.ControlNumero = a.ControlNumero AND aa.Secuencia = a.Secuencia " +
                "LEFT JOIN WH_OrdenCompraDetalle c WITH(NOLOCK)ON b.CompaniaSocio = c.CompaniaSocio AND b.NumeroOrden = c.NumeroOrden AND a.Secuencia = c.Secuencia " +
                "LEFT JOIN WH_OrdenCompra d WITH(NOLOCK) ON b.CompaniaSocio = d.CompaniaSocio AND b.NumeroOrden = d.NumeroOrden " +
                "LEFT JOIN SatelliteCore.dbo.TBMAnalisisAgujas e WITH(NOLOCK) ON e.ControlNumero = a.ControlNumero AND e.ReferenciaSecuencia = a.Linea " +
                "WHERE a.CompaniaSocio = '01000000' AND b.ControlNumero >= '' AND b.NumeroOrden = @NumeroOrden AND LEFT(b.NumeroOrden, 2) <> 'PE' " +
                "ORDER BY a.Linea";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<ListarOrdenCompra>(script, new { NumeroOrden });
            }

            return result;
        }

        public async Task<int> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia)
        {
            int result = 0;

            string script = "SELECT b.ValorEntero3 FROM WH_ControlCalidadDetalle a " +
                "INNER JOIN SatelliteCore.dbo.TBDConfiguracion b ON b.IdConfiguracion = 1 AND b.Grupo = 'RANGO' AND a.CantidadRecibida BETWEEN ValorEntero1 AND ValorEntero2 " +
                "WHERE ControlNumero = @controlNumero AND Linea = @secuencia";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryFirstAsync<int>(script, new{controlNumero, secuencia});
            }
            return result;
        }

        public async Task<string> RegistrarAnalisisAguja (ControlAgujasModel matricula)
        {
            string result;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryFirstAsync<string>("sp_AnalisisAguja_RegistrarAnalisis", matricula, commandType: CommandType.StoredProcedure);
            }
            return result;
        }

        public async Task ValidarLoteCreado(string controlNumero, int secuencia, int codUsuarioSesion)
        {
            using SqlConnection context = new SqlConnection(_appConfig.contextSpring);
            await context.ExecuteAsync("usp_AnalisisAguja_ValidarLoteCreado", new { controlNumero, secuencia, codUsuarioSesion }, commandType: CommandType.StoredProcedure);
        }

        public async Task<(ObtenerAnalisisAgujaModel, List<AnalisisAgujaFlexionEntity>)> AnalisisAgujaFlexion(string loteAnalisis)
        {
            (ObtenerAnalisisAgujaModel cabecera, List<AnalisisAgujaFlexionEntity> detalle) analisis;

            string script = "SELECT ControlNumero, OrdenCompra, Item, DescripcionItem, CodProveedor, Proveedor, CantidadPruebas FROM TBMAnalisisAgujas WHERE Lote = @loteAnalisis " +
                "SELECT IdAnalisis, Lote,TipoRegistro,Llave,Valor,UsuarioRegistro,FechaRegistro FROM TBDAnalisisAgujaFlexion WHERE Lote = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync(script, new { loteAnalisis });
                analisis.cabecera = multi.Read<ObtenerAnalisisAgujaModel>().First();
                analisis.detalle = multi.Read<AnalisisAgujaFlexionEntity>().ToList();
            }

            return analisis;
        }

        public async Task EliminarPruebaFlexionAguja(string loteAnalisis)
        {
            string script = "DELETE TBDAnalisisAgujaFlexion WHERE Lote = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { loteAnalisis });
            }

        }

        public async Task GuardarPruebaFlexionAguja(List<GuardarPruebaFlexionAgujaModel> analisis)
        {
            string script = "INSERT INTO TBDAnalisisAgujaFlexion(Lote, TipoRegistro, Llave, Valor, UsuarioRegistro) VALUES (@lote, @tipoRegistro, @llave, @valor, @usuarioRegistro)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, analisis);
            }
        }

    }
}
