using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class GestionCalidadRepository: IGestionCalidadRepository
    {
        private readonly IAppConfig _appConfig;

        public GestionCalidadRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<List<MateriaPrimaDTO>> ObtenerMateriaPrima(string tipoConsulta, string lote)
        {
            IEnumerable<MateriaPrimaDTO> listaMateriaPrima = new List<MateriaPrimaDTO>();

            string query = "SELECT TH.TipoDocumento, RTRIM(TH.NumeroDocumento) NumeroDocumento,TH.FechaDocumento, RTRIM(TM.DescripcionLocal) Transaccion,RTRIM(TD.Item) Item," +
                "RTRIM(IM.DescripcionLocal)Descripcion,RTRIM(TH.AlmacenCodigo) AlmacenCodigo, TH.ReferenciaTipoDocumento, RTRIM(TH.ReferenciaNumeroDocumento) ReferenciaNumeroDocumento, " +
                "RTRIM(TH.OrdenTrabajo) OrdenFabricacion, TH.Estado, RTRIM(TD.Lote) NumeroAnalisis,TD.Cantidad " +
                "FROM WH_TransaccionHeader TH WITH(NOLOCK) " +
                "INNER JOIN WH_TransaccionDetalle TD WITH(NOLOCK) ON TH.CompaniaSocio = TD.CompaniaSocio AND TH.NumeroDocumento = TD.NumeroDocumento AND TH.TipoDocumento = TD.TipoDocumento " +
                "INNER JOIN WH_ItemMast IM WITH(NOLOCK) ON TD.Item = IM.Item " +
                "INNER JOIN WH_TransaccionMast TM WITH(NOLOCK) ON TH.TransaccionCodigo = TM.TransaccionCodigo AND TH.TipoDocumento = TM.TipoDocumentoGenerado AND TM.estado = 'A' " +
                "WHERE TH.NumeroDocumento >= ' ' AND AlmacenCodigo = 'ALMMAT' AND th.Estado <> 'AN' AND TH.TransaccionCodigo NOT IN('AEG') ";

            if (tipoConsulta == "PT")
                query = $"{query} AND TH.OrdenTrabajo = @lote";
            else
                query = $"{query} AND td.Lote = @lote AND TH.OrdenTrabajo IS NOT NULL";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaMateriaPrima = await context.QueryAsync<MateriaPrimaDTO>(query, new { lote });
            }

            return listaMateriaPrima.ToList();
        }

        public async Task<List<OrdenCompraPorLoteDTO>> OrdenCompraPorlote(List<string> lotes)
        {
            IEnumerable<OrdenCompraPorLoteDTO> listaMateriaPrima = new List<OrdenCompraPorLoteDTO>();

            string query = "SELECT RTRIM(b.NumeroOrden) NumeroOrden, b.ControlNumero, RTRIM(d.NombreCompleto) Proveedor,a.LoteAprobado , RTRIM(ISNULL(e.Item,a.Item)) AS Item, " +
                "RTRIM(ISNULL(e.UnidadCodigo, f.UnidadCodigo)) AS UnidadCodigo, ISNULL(e.Descripcion, f.DescripcionLocal) AS Descripcion, RTRIM(f.NumeroDeParte) NumeroDeParte," +
                "IIF(b.ReferenciaTipo = 'IM', 'IMPORTACION', 'NACIONAL') ReferenciaTipo,a.CantidadRecibida, a.CantidadAceptada, a.CantidadRechazada, a.CantidadTransferida, " +
                "a.FechaAprobacion, a.FechaVencimiento, RTRIM(a.Comentarios) Comentario FROM WH_ControlCalidadDetalle a WITH(NOLOCK) " +
                "INNER JOIN WH_ControlCalidad b WITH(NOLOCK) ON a.CompaniaSocio = b.CompaniaSocio AND a.ControlNumero = b.ControlNumero " +
                "LEFT JOIN WH_OrdenCompra c WITH(NOLOCK) ON c.CompaniaSocio = b.CompaniaSocio AND b.NumeroOrden = c.NumeroOrden " +
                "LEFT JOIN PersonaMast d WITH(NOLOCK) ON d.Persona = c.Proveedor " +
                "LEFT JOIN WH_OrdenCompraDetalle e WITH(NOLOCK) ON b.CompaniaSocio = e.CompaniaSocio AND b.NumeroOrden = e.NumeroOrden AND a.Secuencia = e.Secuencia " +
                "LEFT JOIN WH_ItemMast f WITH(NOLOCK) ON a.Item = f.Item " +
                "WHERE b.ControlNumero >= '' AND b.NumeroOrden >= '' AND b.CompaniaSocio = '01000000' AND a.Estado = 'GE' AND b.ReferenciaTipo IN('OC','IM') " +
                "AND LEFT(b.NumeroOrden, 2) <> 'PE' AND a.LoteAprobado IN @lotes ORDER BY b.CompaniaSocio ASC, b.ControlNumero ASC ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaMateriaPrima = await context.QueryAsync<OrdenCompraPorLoteDTO>(query, new { lotes });
            }

            return listaMateriaPrima.ToList();
        }


        public async Task<List<OrdenFabricacionPorLoteDTO>> OrdenFabricacionPorlotes(List<string> lotes)
        {
            IEnumerable<OrdenFabricacionPorLoteDTO> listaMateriaPrima = new List<OrdenFabricacionPorLoteDTO>();

            string query = "SELECT a.NUMEROLOTE OrdenFabricacion, a.REFERENCIANUMERO Lote,a.PedidoNumero, RTRIM(b.Busqueda) Cliente,a.FechaProduccion,a.FechaExpiracion, " +
                "a.Item,c.NumeroDeParte, c.DescripcionLocal Descripcion, RTRIM(c.UnidadCodigo) UnidadCodigo, a.CantidadProducida,a.Estado, a.DiaProduccion, a.ReferenciaTipo, a.AuditableFlag," +
                "a.TransferidoFlag,a.Comentarios,a.FechaRequerida, (SELECT COUNT(*) FROM EP_Pedido e WITH(NOLOCK) " +
                "INNER JOIN EP_PedidoDetalle f  WITH(NOLOCK) ON e.CompaniaSocio = f.CompaniaSocio AND e.PedidoNumero = f.PedidoNumero WHERE f.CompaniaSocio = '01000000' AND " +
                "f.NumeroLote = a.NumeroLote ) AS MultiPedido FROM EP_PROGRAMACIONLOTE a WITH(NOLOCK) LEFT JOIN WH_ITEMMAST c  WITH(NOLOCK) ON a.ITEM = c.ITEM " +
                "LEFT JOIN PERSONAMAST b  WITH(NOLOCK) ON b.PERSONA = a.CLIENTE WHERE a.COMPANIASOCIO = '01000000' AND a.NumeroLote IN @lotes";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaMateriaPrima = await context.QueryAsync<OrdenFabricacionPorLoteDTO>(query, new { lotes });
            }

            return listaMateriaPrima.ToList();
        }


        public async Task<List<DocumentosPedidosDTO>> OrdenDocumentosPedidosPorLotes(List<string> lotes)
        {
            IEnumerable<DocumentosPedidosDTO> listaMateriaPrima = new List<DocumentosPedidosDTO>();

            string query = "SELECT RTRIM(b.NumeroDocumento) NumeroDocumento,a.TipoDocumento,FormaFacturacion,TipoVenta,FechaDocumento,RTRIM(ClienteRUC)Ruc,RTRIM(ClienteNombre) ClienteNombre," +
                "Clientedireccion,RTRIM(b.ItemSerie) Lote,RTRIM(b.lote) OrdenFabricacion,RTRIM(b.AlmacenCodigo) AlmacenCodigo,RTRIM(ItemCodigo) Item,Descripcion,RTRIM(UnidadCodigo) Unidad," +
                "CantidadEntregada,PrecioUnitario,Monto,MontoTotal FROM CO_Documento a WITH(NOLOCK) INNER JOIN CO_DocumentoDetalle b WITH(NOLOCK) ON a.CompaniaSocio = b.CompaniaSocio AND a.TipoDocumento = b.TipoDocumento " +
                "AND a.NumeroDocumento = b.numerodocumento WHERE b.Lote IN @lotes AND a.Estado <> 'AN'";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaMateriaPrima = await context.QueryAsync<DocumentosPedidosDTO>(query, new { lotes });
            }

            return listaMateriaPrima.ToList();
        }

        public async Task<List<GuiaDocumentos>> OrdenGuiasRelacionadasPorLotes(List<string> lotes)
        {
            IEnumerable<GuiaDocumentos> listaMateriaPrima = new List<GuiaDocumentos>();

            string query = "SELECT a.GuiaNumero,FacturaNumero,a.FechaDocumento,RTRIM(a.ReferenciaNumeroPedido) ReferenciaNumeroPedido,RTRIM(a.DestinatarioRUC) DestinatarioRuc," +
                "RTRIM(a.DestinatarioNombre) DestinatarioNombre,DestinatarioDireccion, RTRIM(b.Lote) Lote, RTRIM(b.ItemCodigo) ItemCodigo,RTRIM(b.Descripcion) Descripcion, " +
                "b.Cantidad,b.CantidadRecibida FROM WH_GuiaRemision a WITH(NOLOCK) INNER JOIN WH_GuiaRemisionDetalle b WITH(NOLOCK) ON a.CompaniaSocio = b.CompaniaSocio " +
                "AND a.SerieNumero = b.SerieNumero AND a.GuiaNumero = b.GuiaNumero WHERE b.Lote IN @lotes AND a.Estado <> 'AN'";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaMateriaPrima = await context.QueryAsync<GuiaDocumentos>(query, new { lotes });
            }

            return listaMateriaPrima.ToList();
        }

        public async Task<List<VentasPorClienteDTO>> VentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            IEnumerable<VentasPorClienteDTO> ventas = new List<VentasPorClienteDTO>();

            string query = "SELECT a.TipoDocumento,RTRIM(a.NumeroDocumento)NumeroDocumento,a.FechaDocumento,a.Estado,a.TipoVenta,MontoTotal,RTRIM(ComercialPedidoNumero) ComercialPedidoNumero, " +
                "RTRIM(Comentarios) Comentarios, RTRIM(b.ItemCodigo) ItemCodigo,RTRIM(c.NumeroDeParte) NumeroDeParte,RTRIM(d.DescripcionLocal) Linea, RTRIM(e.DescripcionLocal) Familia, RTRIM(f.DescripcionLocal) SubFamilia, " +
                "RTRIM(b.Lote) Lote,RTRIM(b.ItemSerie) ItemSerie,b.Descripcion,RTRIM(b.UnidadCodigo) UnidadCodigo,b.CantidadEntregada, b.PrecioUnitario,b.Monto FROM CO_Documento a " +
                "INNER JOIN co_documentodetalle b ON a.CompaniaSocio = b.CompaniaSocio AND a.Tipodocumento = b.TipoDocumento AND a.Numerodocumento = b.Numerodocumento " +
                "INNER JOIN WH_ItemMast c ON b.Itemcodigo = c.Item INNER JOIN WH_ClaseLinea d ON c.Linea = d.Linea INNER JOIN WH_ClaseFamilia e ON e.Linea = c.Linea AND e.Familia = c.Familia " +
                "INNER JOIN wh_clasesubfamilia f ON f.Linea = c.Linea AND f.Familia = c.Familia AND f.SubFamilia = c.SubFamilia " +
                "WHERE a.Estado <> 'AN' AND a.ClienteNumero = @cliente AND a.TipoDocumento <> 'PE' AND a.FechaDocumento BETWEEN @fechaInicio AND @fechaFin";

            if (!string.IsNullOrEmpty(filtros.Linea))
                query = $"{query} AND c.Linea = @linea";

            if (!string.IsNullOrEmpty(filtros.Familia))
                query = $"{query} AND c.Familia = @familia";

            if (!string.IsNullOrEmpty(filtros.SubFamilia))
                query = $"{query} AND c.SubFamilia = @subFamilia";

            if (!string.IsNullOrEmpty(filtros.Item))
                query = $"{query} AND RTRIM(b.ItemCodigo) = @item";

            if (!string.IsNullOrEmpty(filtros.NumeroParte))
                query = $"{query} AND c.NumeroDeParte LIKE @numeroParte";
            

            query = $"{query} ORDER BY a.FechaDocumento";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                ventas = await context.QueryAsync<VentasPorClienteDTO>(query, filtros);
            }

            return ventas.ToList();
        }


        public async Task<IEnumerable<DatosFormatoListarSsomaModel>> ListarSsoma(int TipoDocumento, string Codigo , int Estado)
        {
            IEnumerable<DatosFormatoListarSsomaModel> result = new List<DatosFormatoListarSsomaModel>();

            string query = " SELECT  a.IdSsoma,a.idTipoDocumento ,b.Descripcion TipoDocumentoDescripcion, ISNULL(a.idUbicacionSsoma,0)  idUbicacionSsoma , ISNULL(a.idProteccionSsoma,0) idProteccionSsoma, ISNULL(a.idEstadoSsoma,0)  idEstadoSsoma, ISNULL(c.Descripcion,'') EstadoDescripcion," +
                           " a.CodigoDocumento , a.NombreDocumento , a.FechaPublicacion, a.VersionSsoma , a.Vigencia , ISNULL(a.idSsomaAlmacenamiento, 0) idSsomaAlmacenamiento,a.ArchivoPasivo,a.idSsomaResponsable responsable,a.FechaAprobacion, a.FechaRevision , a.Comentario ,DATEDIFF(DAY,a.FechaPublicacion,a.FechaRevision) Dias " +
                           " FROM dbo.TBMSsoma a "+
                           " INNER JOIN dbo.TBMSsomaTipoDocumento b ON a.idTipoDocumento = b.idTipoDocumento " +
                           " LEFT JOIN dbo.TBMSsomaEstado c ON a.idEstadoSsoma = c.idEstadoSsoma " +
                           " WHERE a.idTipoDocumento = IIF(@TipoDocumento = 0, b.idTipoDocumento, @TipoDocumento) AND " +
                           " a.CodigoDocumento = IIF(@Codigo is null, a.CodigoDocumento, @Codigo)  AND  a.idEstadoSsoma = IIF(@Estado = 0, a.idEstadoSsoma, @Estado)  AND a.Estado = 'A' ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoListarSsomaModel>(query, new { TipoDocumento , Codigo, Estado });
            }

            return result;
        }


        public async Task<object> RegistrarSsoma(DatosFormatoRegistrarSsomaModel dato,string UsuarioSesion)
        {
            dynamic resp = new { mensaje = "", respuesta = false };

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {

                resp = await connection.QuerySingleOrDefaultAsync("usp_Registrar_Editar_Ssoma", new { dato.idSsoma, dato.codigo, dato.nombreDocumento, dato.tipoDocumento, dato.version, dato.vigencia, dato.fechapublicacion, dato.fecharevision, dato.fechaAprobacion, dato.estado, dato.Ubicacion, dato.Almacenamiento, dato.proteccion, dato.responsable, dato.archivopasivo, dato.comentario, UsuarioSesion }, commandType: CommandType.StoredProcedure);
  
                connection.Dispose();
            }

            return resp;
        }

        public async Task<int> EliminarSsoma(int idSsoma, string UsuarioSesion)
        {
            int result = 1;
            string sql = "UPDATE dbo.TBMSsoma SET  UsuarioModificacion=@UsuarioSesion ,FechaModificacion=GETDATE() , Estado='I' WHERE IdSsoma=@idSsoma";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {

                await connection.ExecuteAsync(sql, new { idSsoma, UsuarioSesion });

                connection.Dispose();
            }

            return result;
        }

    }
}
