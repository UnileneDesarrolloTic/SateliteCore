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

            string script = "SELECT Lote, Item, DescripcionItem, OrdenCompra, Proveedor, Cantidad, FechaRegistro, CantidadPruebas FROM TBMAnalisisAgujas AA WITH(NOLOCK) " +
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
                "LEFT JOIN WH_OrdenCompraDetalle c WITH(NOLOCK) ON b.CompaniaSocio = c.CompaniaSocio AND b.NumeroOrden = c.NumeroOrden AND a.Secuencia = c.Secuencia " +
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

            string script = "SELECT IIF(SUBSTRING(b.NumeroDeParte, 12, 1) = '3', CAST(c.ValorDecimal1 AS INT), c.ValorEntero3) Cantidad " +
                "FROM WH_ControlCalidadDetalle a WITH(NOLOCK) INNER JOIN WH_ItemMast b WITH(NOLOCK) ON a.Item = b.Item " +
                "INNER JOIN SatelliteCore.dbo.TBDConfiguracion c WITH(NOLOCK) ON c.IdConfiguracion = 1 AND c.Grupo = 'RANGO' AND a.CantidadRecibida BETWEEN c.ValorEntero1 AND c.ValorEntero2 AND c.Estado = 'A'" +
                "WHERE a.ControlNumero = @controlNumero AND a.Linea = @secuencia";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryFirstOrDefaultAsync<int>(script, new{controlNumero, secuencia});
            }
            return result;
        }

        public async Task<string> RegistrarAnalisisAguja (ControlAgujasModel matricula)
        {
            string result;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryFirstOrDefaultAsync<string>("usp_AnalisisAguja_RegistrarAnalisis", matricula, commandType: CommandType.StoredProcedure);
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

            string script = "SELECT a.ControlNumero, a.OrdenCompra, a.Item, a.DescripcionItem, a.CodProveedor, a.Proveedor, a.CantidadPruebas, " +
                "IIF(SUBSTRING(b.NumeroDeParte, 12, 1) = '3', '300', '400') Serie FROM TBMAnalisisAgujas a WITH(NOLOCK) " +
                "INNER JOIN PROD_UNILENE2.dbo.WH_ItemMast b WITH(NOLOCK) ON a.Item = b.Item WHERE Lote = @loteAnalisis " +
                "SELECT IdAnalisis, Lote,TipoRegistro,Llave,Valor,UsuarioRegistro,FechaRegistro FROM TBDAnalisisAgujaFlexion WITH(NOLOCK) WHERE Lote = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync(script, new { loteAnalisis });
                analisis.cabecera = multi.Read<ObtenerAnalisisAgujaModel>().FirstOrDefault();
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

        public async Task<ObtenerDatosGeneralesDTO> ObtenerDatosGenerales(string loteAnalisis)
        {
            ObtenerDatosGeneralesDTO result = new ObtenerDatosGeneralesDTO();

            string script = "SELECT RTRIM(c.Codigo) CodTipo,RTRIM(c.DescripcionLocal) Tipo,RTRIM(d.Codigo) CodLongitud,RTRIM(d.DescripcionLocal) Longitud,RTRIM(f.Codigo) CodBroca, RTRIM(f.DescripcionLocal) Broca," +
                "RTRIM(g.Codigo) CodAlambre,RTRIM(g.DescripcionLocal) Alambre,IIF(SUBSTRING(b.NumeroDeParte, 12, 1) = '3', '300', '400') Serie,a.OrdenCompra,a.ControlNumero,a.Proveedor,a.Cantidad," +
                "CAST(e.ValorDecimal2 AS INT) UndMuestrear,e.ValorEntero3 UndMuestrearI, CAST(e.ValorDecimal2 AS INT) UndMuestrearIII, a.Observaciones, ISNULL(a.FechaModificacion, a.FechaRegistro) FechaAnalisis " +
                "FROM SatelliteCore.dbo.TBMAnalisisAgujas a WITH(NOLOCK) " +
                    "INNER JOIN WH_ItemMast b WITH(NOLOCK) ON a.Item = b.Item INNER JOIN WH_ItemFormato c ON c.Grupo = '15' AND c.Tabla = '002' AND c.Codigo = SUBSTRING(b.NumeroDeParte, 2, 2) " +
                    "INNER JOIN WH_ItemFormato d WITH(NOLOCK) ON d.Grupo = '15' AND d.Tabla = '003' AND d.Codigo = SUBSTRING(b.NumeroDeParte, 4, 3) " +
                    "INNER JOIN SatelliteCore.dbo.TBDConfiguracion e WITH(NOLOCK) ON e.IDConfiguracion = 3 AND e.Grupo = 'BATCH' AND a.Cantidad BETWEEN e.ValorEntero1 AND e.ValorEntero2 AND e.Estado = 'A'" +
                    "INNER JOIN WH_ItemFormato f WITH(NOLOCK) ON f.Grupo = '15' AND f.Tabla = '004' AND f.Codigo = CASE WHEN LEN(b.NumeroDeParte) IN(11, 14) THEN SUBSTRING(b.NumeroDeParte, 7, 2) ELSE SUBSTRING(b.NumeroDeParte, 7, 3) END " +
                    "INNER JOIN WH_ItemFormato g WITH(NOLOCK) ON g.Grupo = '15' AND g.Tabla = '005' AND g.Codigo = CASE WHEN LEN(b.NumeroDeParte) IN(11, 14) THEN SUBSTRING(b.NumeroDeParte, 9, 3) ELSE SUBSTRING(b.NumeroDeParte, 10, 3) END " +
                    "INNER JOIN SatelliteCore.dbo.TBDConfiguracion h WITH(NOLOCK) ON h.IDConfiguracion = 5 AND h.Grupo = 'TP_Aguja' AND h.ValorTexto1 = SUBSTRING(b.NumeroDeParte, 2, 2) AND h.Estado = 'A'" +
                "WHERE Lote = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryFirstOrDefaultAsync<ObtenerDatosGeneralesDTO>(script, new { loteAnalisis });
            }

            return result;
        }

        public async Task<AnalisisAgujaPlanMuestreoEntity> ObtenerPlanMuestreo(string loteAnalisis)
        {
            AnalisisAgujaPlanMuestreoEntity result = new AnalisisAgujaPlanMuestreoEntity();
            // REVISAR
            string script = "SELECT LoteAnalisis,Cantidad,UndMuestrear,UndMuestrearI,UndMuestrearIII,CajasMuestrear,StatusFlexion,Usuario,Fecha FROM TBDAnalisisAgujaPlanMuestreoFlexion WITH(NOLOCK) " +
                "WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryFirstOrDefaultAsync<AnalisisAgujaPlanMuestreoEntity>(script, new { loteAnalisis });
            }

            return result;
        }

        public async Task EliminarPlanMuestreo(string loteAnalisis)
        {
            string script = "DELETE TBDAnalisisAgujaPlanMuestreoFlexion WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { loteAnalisis });
            }
        }

        public async Task RegistrarPlanMuestreo(AnalisisAgujaPlanMuestreoEntity planMuestreo)
        {
            string script = "INSERT INTO TBDAnalisisAgujaPlanMuestreoFlexion (LoteAnalisis,Cantidad,UndMuestrear,UndMuestrearI,UndMuestrearIII,CajasMuestrear, StatusFlexion, Usuario) " +
                "VALUES (@loteAnalisis,@cantidad,@undMuestrear,@undMuestrearI,@undMuestrearIII,@cajasMuestrear,@statusFlexion,@usuario)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, planMuestreo);
            }
        }

        public async Task<List<AnalisisAgujaPruebaDimensionalEntity>> ObtenerPruebaDimensional(string loteAnalisis)
        {
            List<AnalisisAgujaPruebaDimensionalEntity> result = new List<AnalisisAgujaPruebaDimensionalEntity>();

            string script = "SELECT LoteAnalisis,TipoRegistro,Cantidad,BaseCalculoEstado,Tolerancia,DescripcionAux,CantidadAux,Usuario,Fecha FROM TBDAnalisisAgujaPruebaDimensional WITH(NOLOCK) " +
                "WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = (List<AnalisisAgujaPruebaDimensionalEntity>)await context.QueryAsync<AnalisisAgujaPruebaDimensionalEntity>(script, new { loteAnalisis });
            }

            return result;
        }

        public async Task EliminarPruebaDimensional(string loteAnalisis)
        {
            string script = "DELETE TBDAnalisisAgujaPruebaDimensional WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { loteAnalisis });
            }
        }

        public async Task RegistrarPruebaDimensional(List<AnalisisAgujaPruebaDimensionalEntity> prueba)
        {
            string script = "INSERT INTO TBDAnalisisAgujaPruebaDimensional (LoteAnalisis,TipoRegistro,Cantidad,BaseCalculoEstado,Tolerancia,DescripcionAux,CantidadAux,Usuario) " +
                "VALUES(@loteAnalisis,@tipoRegistro,@cantidad,@baseCalculoEstado,@tolerancia,@descripcionAux,@cantidadAux,@usuario)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, prueba);
            }
        }

        public async Task<IEnumerable<AnalisisAgujaElasticidadPerforacionEntity>> ObtenerPruebaElasticidadPerforacion(string loteAnalisis)
        {
            IEnumerable<AnalisisAgujaElasticidadPerforacionEntity> result = new List<AnalisisAgujaElasticidadPerforacionEntity>();

            string script = "SELECT LoteAnalisis,TipoRegistro,Uno,Dos,Tres,Cuatro,Cinco,Estado,Usuario,Fecha FROM TBDAnalisisAgujaElasticidadPerforacion WITH(NOLOCK) " +
                "WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<AnalisisAgujaElasticidadPerforacionEntity>(script, new { loteAnalisis });
            }

            return result;
        }

        public async Task EliminarPruebaElasticidadPerforacion(string loteAnalisis)
        {
            string script = "DELETE TBDAnalisisAgujaElasticidadPerforacion WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { loteAnalisis });
            }
        }

        public async Task RegistrarPruebaElasticidadPerforacion(List<AnalisisAgujaElasticidadPerforacionEntity> prueba)
        {
            string script = "INSERT INTO TBDAnalisisAgujaElasticidadPerforacion (LoteAnalisis,TipoRegistro,Uno,Dos,Tres,Cuatro,Cinco,Estado,Usuario,Fecha) " +
                "VALUES(@loteAnalisis,@tipoRegistro,@uno,@dos,@tres,@cuatro,@cinco,@estado,@usuario,@fecha)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, prueba);
            }
        }

        public async Task<IEnumerable<AnalisisAgujaPruebaAspectoEntity>> ObtenerPruebaAspecto(string loteAnalisis)
        {
            IEnumerable<AnalisisAgujaPruebaAspectoEntity> result = new List<AnalisisAgujaPruebaAspectoEntity>();

            string script = "SELECT LoteAnalisis,TipoRegistro,Cantidad,BaseCalculoPorcentaje,Tolerancia,Usuario,Fecha FROM TBDAnalisisAgujaPruebaAspecto WITH(NOLOCK) " +
                "WHERE LoteAnalisis = @loteAnalisis ORDER BY TipoRegistro ASC";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<AnalisisAgujaPruebaAspectoEntity>(script, new { loteAnalisis });
            }

            return result;
        }

        public async Task EliminarPruebaAspecto(string loteAnalisis)
        {
            string script = "DELETE TBDAnalisisAgujaPruebaAspecto WHERE LoteAnalisis = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { loteAnalisis });
            }
        }

        public async Task RegistrarPruebaAspecto(PruebaAspectoYObservacionesDTO datos, string loteAnalisis)
        {


            string script = "INSERT INTO TBDAnalisisAgujaPruebaAspecto (LoteAnalisis,TipoRegistro,Cantidad,BaseCalculoPorcentaje,Tolerancia,Usuario,Fecha) " +
                "VALUES(@loteAnalisis,@tipoRegistro,@cantidad,@baseCalculoPorcentaje,@tolerancia,@usuario,@fecha)";

            string actualizarObservacionesScript = "UPDATE TBMAnalisisAgujas SET Observaciones = @observaciones WHERE Lote = @loteAnalisis";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, datos.Pruebas);
                await context.ExecuteAsync(actualizarObservacionesScript, new { observaciones = datos.Observaciones, loteAnalisis });
            }
        }

    }
}
