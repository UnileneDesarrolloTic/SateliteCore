using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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

        public async Task<DatosFormatoPlanOrdenServicosD> ObtenerNumeroGuias(string numeroguia)
        {
            DatosFormatoPlanOrdenServicosD result = new DatosFormatoPlanOrdenServicosD();

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await context.QueryFirstOrDefaultAsync<DatosFormatoPlanOrdenServicosD>("usp_informacion_retorno_guia_satelite", new { numeroguia }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<int> RegistrarRetornoGuia(DatosFormatoRetornoGuiaRequest dato)
        {
            string script = "UPDATE TLOG_PLAN_ORDEN_SERVICIO_D SET FECHA_RETORNO=GETDATE()  WHERE " +
                            "CONCAT(RTRIM(SERIE),'-',RTRIM(NUMERO_DOCUMENTO))=@numeroGuia AND Estado='A'";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(script, dato);
            }

            return 1;
        }

        public async Task<IEnumerable<DatosFormatosReporteRetornoGuia>> ListarRetornoGuia(DatosFormatoGestionGuiasClienteModel dato)
        {
            IEnumerable<DatosFormatosReporteRetornoGuia> result = new List<DatosFormatosReporteRetornoGuia>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatosReporteRetornoGuia>("usp_reporte_retornoguia_excel", new { dato.idCliente, dato.destino, dato.transportista, dato.fechaInicio, dato.fechaFin }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato)
        {

            IEnumerable<DatosFormatoItemVentas> result = new List<DatosFormatoItemVentas>();
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoItemVentas>("usp_Lista_Item_Ventas", new { dato.Item, dato.Codsut, dato.Descripcion, dato.Origen, dato.idmarca, dato.idAgrupador, dato.idSubAgrupador, dato.idLinea, dato.idfamilia, dato.idSubFamilia }, commandType: CommandType.StoredProcedure);
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item)
        {
            IEnumerable<DatosFormatoItemLoteAlmacen> result = new List<DatosFormatoItemLoteAlmacen>();
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoItemLoteAlmacen>("usp_Buscar_Item_Ventas", new { Item }, commandType: CommandType.StoredProcedure);
            }
            return result;
        }


        public async Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle()
        {
            IEnumerable<DatosFormatoDetalledelItemVentas> result = new List<DatosFormatoDetalledelItemVentas>();
            string script = "SELECT RTRIM(b.Item) Item, RTRIM(d.DescripcionLocal) DescripcionItem, RTRIM(b.AlmacenCodigo) AlmacenCodigo,RTRIM(c.DescripcionLocal) DescripcionLocal, RTRIM(b.Lote) Lote , ISNULL(SUM(b.StockActual), 0)  StockActual, ISNULL(SUM(b.StockComprometido), 0) StockComprometido , " +
                            "ISNULL(SUM(b.StockActual), 0) - ISNULL(SUM(b.StockComprometido), 0) StockDisponible FROM[PROD_UNILENE2]..WH_ITEMMAST  a " +
                            "LEFT JOIN [PROD_UNILENE2]..WH_ItemAlmacenLote b ON a.Item = b.Item " +
                            "INNER JOIN [PROD_UNILENE2]..WH_AlmacenMast c ON c.AlmacenCodigo = b.AlmacenCodigo " +
                            "INNER JOIN [PROD_UNILENE2]..WH_ITEMMAST d ON b.Item = d.Item " +
                            "WHERE b.AlmacenCodigo IN(SELECT AlmacenCodigo FROM[PROD_UNILENE2]..WH_AlmacenMast WHERE Estado = 'A' AND AlmacenVentaFlag = 'S') " +
                            "AND b.StockActual > 0 AND d.Linea IN ('P','D') AND d.Estado='A'" +
                            "GROUP BY b.AlmacenCodigo,c.DescripcionLocal,d.DescripcionLocal,b.Item ,b.Lote";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoDetalledelItemVentas>(script);
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato)
        {
            IEnumerable<DatosFormatoDetalleComprometidoItem> result = new List<DatosFormatoDetalleComprometidoItem>();
            string script = "SELECT a.CompaniaSocio, a.TipoDocumento, RTRIM(a.NumeroDocumento) NumeroDocumento, RTRIM(a.ClienteNombre) ClienteNombre, a.FechaDocumento, a.FechaVencimiento, " +
                            "a.TipoVenta , a.Vendedor, RTRIM(a.comentarios) comentarios ,b.TipoDetalle, b.ItemCodigo, b.Descripcion, RTRIM(b.UnidadCodigo) UnidadCodigo, RTRIM(a.AlmacenCodigo) AlmacenCodigo , " +
                            "RTRIM(c.Busqueda) Busqueda, RTRIM(d.DescripcionLocal) DescripcionLocal, SUM(b.CantidadPedida - CantidadEntregada) as CantidadPedida " +
                            " FROM [PROD_UNILENE2]..CO_Documento a WITH(NOLOCK) " +
                            "INNER JOIN [PROD_UNILENE2]..CO_DocumentoDetalle b WITH(NOLOCK)" +
                            "ON (a.CompaniaSocio = b.CompaniaSocio and a.TipoDocumento = b.TipoDocumento and a.NumeroDocumento = b.NumeroDocumento) LEFT JOIN[PROD_UNILENE2]..PersonaMast c WITH(NOLOCK)" +
                            "ON (a.Vendedor = c.Persona)" +
                            "INNER JOIN[PROD_UNILENE2]..CO_TipoDocumento d ON a.TipoDocumento = d.TipoDocumento " +
                            "WHERE(a.CompaniaSocio = '01000000') AND((a.TipoDocumento = 'PE' AND a.Estado = 'AP') " +
                            "OR(a.TipoDocumento in ('FC', 'BV', 'PK') AND a.Estado <> 'AN')) AND(b.AlmacenCodigo = @almacenCodigo) AND(b.TipoDetalle = 'I') AND " +
                            "(b.Estado <> 'CE') AND(b.ItemCodigo = @item) AND(b.Lote = @lote) AND(b.CantidadPedida > IsNull(b.CantidadEntregada, 0)) " +
                            "GROUP BY a.CompaniaSocio, a.TipoDocumento, a.NumeroDocumento, a.ClienteNombre, a.FechaDocumento, a.FechaVencimiento, " +
                            "a.TipoVenta , a.Vendedor, a.comentarios , a.ClienteNombre, b.TipoDetalle, b.ItemCodigo, b.Descripcion, b.UnidadCodigo, a.AlmacenCodigo , c.Busqueda, d.DescripcionLocal";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoDetalleComprometidoItem>(script, new { dato.item, dato.lote, dato.almacenCodigo });
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo)
        {
            IEnumerable<DatosFormatoMateriaPrimaItemLogistica> result = new List<DatosFormatoMateriaPrimaItemLogistica>();

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await context.QueryAsync<DatosFormatoMateriaPrimaItemLogistica>("SP_UNILENE_CO_LISTAR_DETALLE_COTIZACION", new { NumeroDocumento, Tipo }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoDetalleRecetaMPLogistica>> BuscardDetalleRecetaMP(string Item, string Cantidad)
        {
            IEnumerable<DatosFormatoDetalleRecetaMPLogistica> result = new List<DatosFormatoDetalleRecetaMPLogistica>();

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await context.QueryAsync<DatosFormatoDetalleRecetaMPLogistica>("SP_UNILENE_CO_LISTAR_DETALLE_RECETA", new { Item, Cantidad }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }
    }
}
