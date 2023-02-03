using Dapper;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request.GestionCalidad;
using SatelliteCore.Api.Models.Request.GestorDocumentario;
using SatelliteCore.Api.Models.Response.GestionCalidad;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class GestionCalidadRepository : IGestionCalidadRepository
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

            string query = "SELECT a.TipoDocumento,RTRIM(a.NumeroDocumento)NumeroDocumento,a.FechaDocumento,a.Estado,RTRIM(a.ClienteNombre) Cliente,a.TipoVenta,MontoTotal,RTRIM(ComercialPedidoNumero) ComercialPedidoNumero, " +
                "RTRIM(Comentarios) Comentarios, RTRIM(b.ItemCodigo) ItemCodigo,RTRIM(c.NumeroDeParte) NumeroDeParte,RTRIM(d.DescripcionLocal) Linea, RTRIM(e.DescripcionLocal) Familia, RTRIM(f.DescripcionLocal) SubFamilia, " +
                "RTRIM(b.Lote) Lote,RTRIM(b.ItemSerie) ItemSerie,b.Descripcion,RTRIM(b.UnidadCodigo) UnidadCodigo,b.CantidadEntregada, b.PrecioUnitario,b.Monto FROM CO_Documento a " +
                "INNER JOIN co_documentodetalle b ON a.CompaniaSocio = b.CompaniaSocio AND a.Tipodocumento = b.TipoDocumento AND a.Numerodocumento = b.Numerodocumento " +
                "INNER JOIN WH_ItemMast c ON b.Itemcodigo = c.Item INNER JOIN WH_ClaseLinea d ON c.Linea = d.Linea INNER JOIN WH_ClaseFamilia e ON e.Linea = c.Linea AND e.Familia = c.Familia " +
                "INNER JOIN wh_clasesubfamilia f ON f.Linea = c.Linea AND f.Familia = c.Familia AND f.SubFamilia = c.SubFamilia " +
                "WHERE a.Estado <> 'AN' AND a.TipoDocumento <> 'PE'";

            if (!string.IsNullOrEmpty(filtros.Lote))
                query = $"{query} AND b.Lote = @lote";

            if (filtros.Cliente != 0)
                query = $"{query} AND a.ClienteNumero = @cliente";

            if (filtros.FechaInicio != null && filtros.FechaFin != null)
                query = $"{query} AND a.FechaDocumento BETWEEN @fechaInicio AND @fechaFin";

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


        public async Task<(List<ListaReclamosDTO>, int cantidadRegistros)> ListarReclamosQuejas(FiltrosListaReclamosDTO filtros)
        {
            IEnumerable<ListaReclamosDTO> reclamos = new List<ListaReclamosDTO>();
            int registros = 0;

            string query = "IF OBJECT_ID('Tempdb..#Temp_Paginacion') IS NOT NULL DROP TABLE #Temp_Paginacion " +
                "SELECT a.CodReclamo, RTRIM(b.NombreCompleto) NombreCliente, RTRIM(d.DescripcionLocal) TipoCliente, " +
                "RTRIM(e.DescripcionCorta) Nacionalidad, IIF(c.Nacionalidad = 'E', 'Extranjera', 'Nacional') Territorio, " +
                "CASE a.Estado WHEN 'P' THEN 'PROCESO' WHEN 'C' THEN 'COMPLETADO' END Estado, a.FechaRegistro, a.UsuarioRegistro, " +
                "DATEDIFF(DAY, a.FechaRegistro, GETDATE()) DiferenciaDias " +
                "INTO #Temp_Paginacion FROM SatelliteCore.dbo.TBMReclamos a WITH(NOLOCK) INNER JOIN PersonaMast b WITH(NOLOCK) ON a.Cliente = b.Persona " +
                "INNER JOIN ClienteMast c WITH(NOLOCK) ON b.Persona = c.Cliente INNER JOIN CO_TipoCliente d WITH(NOLOCK) ON c.TipoCliente = d.TipoCliente " +
                "LEFT JOIN Pais e WITH(NOLOCK) ON b.Nacionalidad = e.Pais WHERE CAST(a.FechaRegistro AS DATE) BETWEEN @fechaInicio AND @fechaFin ";

            if (filtros.Cliente > 0)
                query = $"{query} AND a.Cliente = @Cliente";

            if (!string.IsNullOrEmpty(filtros.CodReclamo))
                query = $"{query} AND a.codReclamo = @codReclamo";

            if (!string.IsNullOrEmpty(filtros.Territorio))
                query = $"{query} AND c.Nacionalidad = @territorio";

            query = $"{query} SELECT * FROM #Temp_Paginacion ORDER BY DiferenciaDias DESC OFFSET (@pagina - 1) * @registrosPorPagina " +
                $"ROWS FETCH NEXT @registrosPorPagina ROWS ONLY; SELECT COUNT(1) CantidadRegistros FROM #Temp_Paginacion";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                reclamos = await context.QueryAsync<ListaReclamosDTO>(query, filtros);
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync(query, filtros);
                reclamos = multi.Read<ListaReclamosDTO>().ToList();
                registros = multi.Read<int>().FirstOrDefault();
            }

            return (reclamos.ToList(), registros);

        }

        public async Task<IEnumerable<DatosFormatoListarSsomaModel>> ListarSsoma(int tipoDocumento, string codigo , int estado)
        {
            IEnumerable<DatosFormatoListarSsomaModel> result = new List<DatosFormatoListarSsomaModel>();

            string query = " SELECT  a.IdSsoma,a.idTipoDocumento ,b.Descripcion TipoDocumentoDescripcion, ISNULL(a.idUbicacionSsoma,0)  idUbicacionSsoma , ISNULL(a.idProteccionSsoma,0) idProteccionSsoma, ISNULL(a.idEstadoSsoma,0)  idEstadoSsoma, ISNULL(c.Descripcion,'') EstadoDescripcion," +
                           " a.CodigoDocumento , a.NombreDocumento , a.FechaPublicacion, a.VersionSsoma , a.Vigencia , ISNULL(a.idSsomaAlmacenamiento, 0) idSsomaAlmacenamiento,a.ArchivoPasivo,a.idSsomaResponsable responsable,a.FechaAprobacion, a.FechaRevision , a.Comentario ,DATEDIFF(DAY,a.FechaCreacion,a.FechaRevision) Dias " +
                           " FROM dbo.TBMSsoma a "+
                           " INNER JOIN dbo.TBMSsomaTipoDocumento b ON a.idTipoDocumento = b.idTipoDocumento " +
                           " LEFT JOIN dbo.TBMSsomaEstado c ON a.idEstadoSsoma = c.idEstadoSsoma " +
                           " WHERE a.Estado = 'A'";

            if(tipoDocumento != 0)
                query = $"{query} AND a.idTipoDocumento = @tipoDocumento";
            
            if(!string.IsNullOrEmpty(codigo))
                query = $"{query} AND a.CodigoDocumento = @codigo";

            if (estado != 0)
                query = $"{query} AND a.idEstadoSsoma = @estado";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoListarSsomaModel>(query, new { tipoDocumento , codigo, estado });
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

        public async Task RegistrarReclamo(TBMReclamosEntity reclamo)
        {
            string query = "INSERT INTO TBMReclamos(CodReclamo, Cliente, Estado, UsuarioRegistro, FechaRegistro) VALUES (@codReclamo, @cliente, @estado, @usuarioRegistro, @fechaRegistro)";

            using SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB);
            await context.ExecuteAsync(query, reclamo);
        }

        public async Task<CabeceraDetalleReclamoDTO> ObtenerCabeceraDetalleReclamo(string codigoReclamo)
        {

            CabeceraDetalleReclamoDTO cabeceraReclamo = new CabeceraDetalleReclamoDTO();
            string query = "SELECT a.Cliente, a.FechaRegistro, RTRIM(b.NombreCompleto) RazonSocial, RTRIM(b.Documento) Documento," +
                "RTRIM(c.DescripcionCorta) Pais,  IIF(d.Nacionalidad = 'E', 'Extranjera', 'Nacional') Territorio, a.Estado " +
                "FROM TBMReclamos a LEFT JOIN PROD_UNILENE2.dbo.PersonaMast b ON a.Cliente = b.Persona " +
                "INNER JOIN PROD_UNILENE2.dbo.ClienteMast d ON b.Persona = d.Cliente " +
                "LEFT JOIN PROD_UNILENE2.dbo.Pais c ON b.Nacionalidad = c.Pais WHERE CodReclamo = @codigoReclamo";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                cabeceraReclamo = await context.QueryFirstOrDefaultAsync<CabeceraDetalleReclamoDTO>(query, new { codigoReclamo });
            }

            return cabeceraReclamo;
        }

        public async Task<IEnumerable<LotesFiltradosReclamo>> LotesFiltradosReclamo(FiltrosLotesReclamosDTO filtros)
        {
            IEnumerable<LotesFiltradosReclamo> lotes = new List<LotesFiltradosReclamo>();

            string query = "SELECT do.FechaDocumento, RTRIM(td.DescripcionLocal) TipoDocumento, RTRIM(do.TipoDocumento) CodTipoDocumento, RTRIM(do.NumeroDocumento) NumeroDocumento, RTRIM(dd.ItemSerie) Lote, RTRIM(dd.Lote) OrdenFabricacion, RTRIM(dd.ItemCodigo) Item, dd.Descripcion," +
                "dd.CantidadPedida, dd.CantidadEntregada, dd.PrecioUnitario, dd.Monto MontoTotal FROM CO_DocumentoDetalle dd WITH(NOLOCK)" +
                "INNER JOIN CO_Documento do WITH(NOLOCK) ON dd.CompaniaSocio = do.CompaniaSocio AND dd.tipodocumento = do.tipodocumento AND dd.numerodocumento = do.numerodocumento " +
                "INNER JOIN CO_TipoDocumento td ON dd.TipoDocumento = td.TipoDocumento " +
                "WHERE dd.CompaniaSocio = '01000000' AND dd.TipoDocumento <> 'PE' AND ClienteNumero = @cliente AND do.Estado <> 'AN' AND CantidadPedida > 0";

            if (filtros.TipoFiltro == "L")
                query = $"{query} AND dd.ItemSerie = @valorFiltro";

            if (filtros.TipoFiltro == "O")
                query = $"{query} AND dd.Lote = @valorFiltro";

            query = $"{query} ORDER BY do.FechaDocumento DESC";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                lotes = await context.QueryAsync<LotesFiltradosReclamo>(query, filtros);
            }

            return lotes;
        }

        public async Task<DatosLoteReclamoDTO> DatosItemLote(string lote)
        {
            DatosLoteReclamoDTO datos = new DatosLoteReclamoDTO();

            string query = "SELECT a.Item, b.DescripcionLocal, RTRIM(ISNULL(c.DescripcionLocal, b.Linea)) Linea," +
                "RTRIM(ISNULL(d.DescripcionLocal, b.Familia)) Familia, RTRIM(ISNULL(e.DescripcionLocal, b.SubFamilia)) SubFamilia," +
                "RTRIM(ISNULL(f.DescripcionLocal, LEFT(b.NumeroDeParte, 2))) Marca FROM ( SELECT TOP 1 RTRIM(Item) Item " +
                "FROM WH_ItemAlmacenLote WHERE LoteFabricacion = @lote AND AlmacenCodigo IN('ALMCMPT', 'ALMDRO', 'ALMLICIT', 'ALMVTO') " +
                ") a INNER JOIN WH_ItemMast b ON a.Item = b.Item LEFT JOIN WH_ClaseLinea c ON b.Linea = c.Linea " +
                "LEFT JOIN WH_ClaseFamilia d ON b.Linea = d.Linea AND b.Familia = d.Familia " +
                "LEFT JOIN WH_ClaseSubFamilia e ON b.Linea = e.Linea AND b.Familia = e.Familia AND b.SubFamilia = e.SubFamilia " +
                "LEFT JOIN WH_Marcas f ON f.MarcaCodigo = LEFT(b.NumeroDeParte, 2)";


            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                datos = await context.QueryFirstOrDefaultAsync<DatosLoteReclamoDTO>(query, new { lote });
            }

            return datos;

        }

        public async Task<int> GuardarDetalleReclamo(TBDReclamosEntity detalle)
        {
            int idDetalle = 0;

            string query = "INSERT INTO TBDReclamos(CodReclamo,Lote,OrdenFabricacion,TipoDocumento,Documento,Item,Motivo,Remitente,Reingreso,PorCarta,Solicitud," +
                "FechaIncidencia,Clasificacion,AreaInvolucrada,Cantidad,Observaciones,Estado,UsuarioRegistro,FechaRegistro)" +
                "VALUES(@CodReclamo,@Lote,@OrdenFabricacion,@TipoDocumento,@Documento,@Item,@Motivo,@Remitente,@Reingreso,@PorCarta,@Solicitud,@FechaIncidencia," +
                "@Clasificacion,@AreaInvolucrada,@Cantidad,@Observaciones,@Estado,@UsuarioRegistro,@FechaRegistro) " +
                "SELECT SCOPE_IDENTITY() IdDetalle";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                idDetalle = await context.QueryFirstOrDefaultAsync<int>(query, detalle);
            }

            return idDetalle;
        }

        public async Task<int> ActualizarDetalleLoteReclamo(TBDReclamosEntity detalle)
        {
            int registrosActualizados = 0;

            string query = "UPDATE TBDReclamos SET Motivo = @motivo, Solicitud = @solicitud, FechaIncidencia = @fechaIncidencia, " +
                "Clasificacion = @clasificacion, AreaInvolucrada = @areaInvolucrada, Cantidad = @cantidad, Observaciones = @observaciones, " +
                "Remitente = @remitente, PorCarta=@PorCarta, Reingreso=@Reingreso " +
                "WHERE CodReclamo = @codReclamo AND Lote = @lote AND OrdenFabricacion = @ordenFabricacion " +
                    "AND TipoDocumento = @tipoDocumento AND Documento = @documento ";

            using SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB);
            {
                registrosActualizados = await context.ExecuteAsync(query, detalle);
            }

            return registrosActualizados;
        }

        public async Task<bool> ValidarExisteDetalleReclamo(string codReclamo, string lote, string ordenFabricacion, string tipoDocumento, string documento)
        {
            bool validacion = false;

            string query = "SELECT 1 FROM TBDReclamos WHERE CodReclamo = @codReclamo AND Lote = @lote AND OrdenFabricacion = @ordenFabricacion " +
                "AND TipoDocumento = @tipoDocumento AND Documento = @documento";

            query = QueryScript.ConsultaValidacion(query);

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                validacion = await context.QueryFirstOrDefaultAsync<bool>(query, new { codReclamo, lote, ordenFabricacion, tipoDocumento, documento});
            }

            return validacion;
        
        
        }

        public async Task<string> ObtenerEstadoLoteDetalleReclamo (int idDetalle)
        {
            string estado = "";

            string query = "SELECT Estado FROM TBDReclamos WHERE IdDetalle = @idDetalle ";

            query = QueryScript.ConsultaValidacion(query);

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                estado = await context.QueryFirstOrDefaultAsync<string>(query, new { idDetalle });
            }

            return estado;

        }



        public async Task<List<DetalleReclamoDTO>> ListarDetalleReclamo(string codReclamo)
        {
            IEnumerable<DetalleReclamoDTO> detalles = new List<DetalleReclamoDTO>();

            string query = "SELECT a.Documento, a.Lote, a.OrdenFabricacion, RTRIM(ISNULL(c.DescripcionLocal, b.Linea)) Linea, RTRIM(ISNULL(d.DescripcionLocal, b.Familia)) Familia, " +
                "a.Item, b.DescripcionLocal DescripcionItem, RTRIM(ISNULL(f.DescripcionLocal, LEFT(b.NumeroDeParte, 2))) Marca, a.Cantidad, " +
                "CASE a.Estado WHEN 'P' THEN 'PROCESO' WHEN'A' THEN 'ACEPTADO' WHEN 'R' THEN 'RECHAZADO' END Estado, a.UsuarioRegistro, a.FechaRegistro " +
                "FROM SatelliteCore.dbo.TBDReclamos a WITH(NOLOCK) " +
                "INNER JOIN WH_ItemMast b WITH(NOLOCK) ON a.Item = b.Item LEFT JOIN WH_ClaseLinea c WITH(NOLOCK) ON b.Linea = c.Linea " +
                "LEFT JOIN WH_ClaseFamilia d WITH(NOLOCK) ON b.Linea = d.Linea AND b.Familia = d.Familia " +
                "LEFT JOIN WH_Marcas f WITH(NOLOCK) ON f.MarcaCodigo = LEFT(b.NumeroDeParte, 2) WHERE a.CodReclamo = @codReclamo ORDER BY a.IdDetalle";

            using (SqlConnection context =  new SqlConnection(_appConfig.contextSpring))
            {
                detalles = await context.QueryAsync<DetalleReclamoDTO>(query, new { codReclamo });
            }

            return detalles.ToList();
        }

        public async Task<CabeceraReclamoLoteDTO> CabeceraReclamoLote (string codReclamo, string lote, string documento)
        {
            CabeceraReclamoLoteDTO cabecera = new CabeceraReclamoLoteDTO();

            string query = "SELECT a.IdDetalle Id, a.OrdenFabricacion,b.FechaDocumento,a.TipoDocumento CodTipoDocumento, ISNULL(RTRIM(i.DescripcionLocal), a.TipoDocumento) TipoDocumento,a.Documento,a.Item," +
                "c.DescripcionLocal Descripcion,RTRIM(ISNULL(d.DescripcionLocal, c.Linea)) Linea,RTRIM(ISNULL(e.DescripcionLocal,c.Familia)) Familia," +
                "RTRIM(ISNULL(f.DescripcionLocal,c.SubFamilia)) SubFamilia,RTRIM(ISNULL(g.DescripcionLocal,LEFT(c.NumeroDeParte,2))) Marca, h.CantidadPedida, h.CantidadEntregada," +
                "h.PrecioUnitario,h.Monto,a.Motivo,a.Solicitud,a.Remitente,a.Reingreso,a.PorCarta,a.FechaIncidencia,a.Clasificacion,a.AreaInvolucrada,a.Cantidad,a.Observaciones,a.Estado," +
                "a.TipoEnvioRespuesta,a.DestinatarioRespuesta,a.LoteCanjeRespuesta,a.Respuesta,a.UsuarioRespuesta,a.FechaRespuesta " +
                "FROM SatelliteCore.dbo.TBDReclamos a WITH(NOLOCK) INNER JOIN CO_DocumentoDetalle h WITH(NOLOCK) ON h.CompaniaSocio = '01000000' AND a.TipoDocumento = h.TipoDocumento " +
                "AND a.Documento = h.NumeroDocumento AND a.Item = h.ItemCodigo INNER JOIN CO_Documento b ON b.CompaniaSocio = h.CompaniaSocio " +
                "AND b.TipoDocumento = a.TipoDocumento AND b.NumeroDocumento = a.Documento INNER JOIN WH_ItemMast c ON c.Item = a.Item LEFT JOIN WH_ClaseLinea d ON c.Linea = d.Linea " +
                "LEFT JOIN WH_ClaseFamilia e WITH(NOLOCK) ON c.Linea = e.Linea AND c.Familia = e.Familia LEFT JOIN WH_claseSubFamilia f ON f.Linea = c.Linea AND f.Familia = c.Familia " +
                "AND f.SubFamilia = c.SubFamilia LEFT JOIN WH_Marcas g WITH(NOLOCK) ON g.MarcaCodigo = LEFT(c.NumeroDeParte, 2) " +
                "LEFT JOIN CO_TipoDocumento i WITH(NOLOCK) ON i.TipoDocumento = a.TipoDocumento " +
                "WHERE CodReclamo = @codReclamo AND a.Lote = @lote AND Documento = @documento";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                cabecera = await context.QueryFirstOrDefaultAsync<CabeceraReclamoLoteDTO>(query, new { codReclamo, lote, documento });
            }

            return cabecera;
        }

        public async Task<int> RegistrarRespuestaReclamo(RespuestaReclamoDTO respuesta)
        {
            int registrosModificados = 0;

            string query = "UPDATE TBDReclamos SET Estado = @estado, TipoEnvioRespuesta = @tipoEnvio, DestinatarioRespuesta = @destinatario," +
                "LoteCanjeRespuesta = @loteCanje, Respuesta = @respuesta, UsuarioRespuesta = @usuario, FechaRespuesta = GETDATE() " +
                "WHERE IdDetalle = @idDetalle";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                registrosModificados = await context.ExecuteAsync(query, respuesta);
            }

            return registrosModificados;
        }

        public async Task Validar_ActualizarEstadoReclamo(int idDetalle)
        {
            string query = "DECLARE @CodReclamo VARCHAR(10); SELECT @CodReclamo = CodReclamo FROM TBDReclamos WHERE IdDetalle = @idDetalle " +
                "IF NOT EXISTS(SELECT 1 FROM TBDReclamos WHERE CodReclamo = @CodReclamo AND Estado = 'P') BEGIN UPDATE TBMReclamos SET Estado = 'C' WHERE CodReclamo = @CodReclamo END";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(query, new { idDetalle });
            }

        }
    }
}
