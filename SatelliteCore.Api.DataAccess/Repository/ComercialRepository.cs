using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Report.Comercial;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Comercial;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ComercialRepository : IComercialRepository
    {
        private readonly IAppConfig _appConfig;

        public ComercialRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }
        public async Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos)
        {
            (List<CotizacionEntity> ListaCertificados, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_COT_ConsultarCotizaciones", datos, commandType: CommandType.StoredProcedure))
                {
                    result.ListaCertificados = result_db.Read<CotizacionEntity>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }

                connection.Dispose();
            }

            return result;
        }

        public async Task<FormatoCotizacionEntity> ObtenerEstructuraFormato(DatosEstructuraFormatoCotizacion datos)
        {
            FormatoCotizacionEntity result = new FormatoCotizacionEntity();
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_ObtenerEstructuraFormatoCotizacion", new { datos.IdFormato, datos.NumeroDocumento }, commandType: CommandType.StoredProcedure))
                {
                    result.Campos = result_db.Read<FormatoCotizacionEntity.Campo>().ToList();
                    result.IdFormato = result_db.Read<int>().First();
                    result.NombreFormato = "";
                }

                using (var result_db = await connection.QueryMultipleAsync("usp_ObtenerEstructuraFormatoCotizacionDetalle", new { datos.IdFormato, datos.NumeroDocumento }, commandType: CommandType.StoredProcedure))
                {
                    result.Cabeceras = result_db.Read<FormatoCotizacionEntity.Cabecera>().ToList();
                    var Lista = result_db.Read<object>().ToList();
                    int cnt = result.Cabeceras.Count;
                    foreach (var item in Lista)
                    {
                        FormatoCotizacionEntity.Filas filas = new FormatoCotizacionEntity.Filas();
                        var Heading = ((IDictionary<string, object>)item).Keys.ToArray();
                        var tmp = ((IDictionary<string, object>)item);

                        for (int i = 1; i <= cnt; i++)
                        {
                            FormatoCotizacionEntity.Filas.Fila fila = new FormatoCotizacionEntity.Filas.Fila
                            {
                                ValorColumna = tmp[Heading[i]].ToString()
                            };
                            filas.lstFilas.Add(fila);
                        }
                        result.Data.Add(filas);
                    }
                }
                connection.Dispose();
            }
            return result;
        }

        public async Task<int> RegistrarRespuestas(FormatoCotizacionRespuesta datos)
        {
            int IdCotizacionGenerada;

            try
            {
                using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
                {
                    using (var result_db = await connection.QueryMultipleAsync("usp_RegistrarCotizacionGenerada", new { datos.IdFormato, datos.NroDocumento, datos.CodUsuario }, commandType: CommandType.StoredProcedure))
                    {
                        IdCotizacionGenerada = result_db.Read<int>().First();
                    }
                    foreach (var item in datos.Campos)
                    {
                        DynamicParameters parameter = new DynamicParameters();
                        parameter.Add("@IdCotizacionGenerada", IdCotizacionGenerada, DbType.Int32, ParameterDirection.Input);
                        parameter.Add("@IdCampo", item.IdCampo, DbType.Int32, ParameterDirection.Input);
                        parameter.Add("@DescripcionCampo", item.DescripcionCampo, DbType.String, ParameterDirection.Input);
                        parameter.Add("@CodigoDescripcionCampo", item.CodigoDescripcionCampo, DbType.String, ParameterDirection.Input);
                        parameter.Add("@TipoCampo", item.TipoCampo, DbType.String, ParameterDirection.Input);
                        parameter.Add("@Respuesta", item.Respuesta, DbType.String, ParameterDirection.Input);
                        parameter.Add("@CodigoRespuesta", item.CodigoRespuesta, DbType.String, ParameterDirection.Input);

                        using (var result_db = await connection.QueryMultipleAsync("usp_RegistrarRespuestasCotizaciones", parameter, commandType: CommandType.StoredProcedure))
                        {

                        }
                    }
                    DynamicParameters parameter2 = new DynamicParameters();
                    parameter2.Add("@IdCotizacionGenerada", IdCotizacionGenerada, DbType.Int32, ParameterDirection.Input);
                    parameter2.Add("@IdCampo", 9999, DbType.Int32, ParameterDirection.Input);
                    parameter2.Add("@DescripcionCampo", "Detalle", DbType.String, ParameterDirection.Input);
                    parameter2.Add("@CodigoDescripcionCampo", "", DbType.String, ParameterDirection.Input);
                    parameter2.Add("@TipoCampo", "T", DbType.String, ParameterDirection.Input);
                    parameter2.Add("@Respuesta", datos.Detalle, DbType.String, ParameterDirection.Input);
                    parameter2.Add("@CodigoRespuesta", "", DbType.String, ParameterDirection.Input);
                    using (var result_db = await connection.QueryMultipleAsync("usp_RegistrarRespuestasCotizaciones", parameter2, commandType: CommandType.StoredProcedure))
                    {

                    }

                    connection.Dispose();
                }
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public async Task<List<DetalleProtocoloAnalisis>> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos)
        {
            IEnumerable<DetalleProtocoloAnalisis> listaProtocoloAnalisis = new List<DetalleProtocoloAnalisis>();

            using (SqlConnection connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                listaProtocoloAnalisis = await connection.QueryAsync<DetalleProtocoloAnalisis>("usp_ListarProtocoloAnalisis", datos, commandType: CommandType.StoredProcedure);
            }

            return listaProtocoloAnalisis.ToList();
        }

        public async Task<List<DetalleProtocoloAnalisis>> ListaProtocolosPorFacturaOPedido(DatosProtocoloAnalisisListado datos)
        {
            IEnumerable<DetalleProtocoloAnalisis> listaProtocolo = new List<DetalleProtocoloAnalisis>();
            string query = "SELECT	RTRIM(a.NumeroDocumento) NumeroDocumento, a.FechaDocumento, a.FechaVencimiento, RTRIM(a.ClienteNombre) ClienteNombre, " +
                    "RTRIM(b.ItemCodigo) ItemCodigo, b.Descripcion, b.CantidadPedida, RTRIM(b.ItemSerie) AS Lote, f.FechaVencimiento AS FechaExpiracion, " +
                    "RTRIM(b.Lote) AS OrdenFabricacion, RTRIM(a.Comentarios) Comentarios " +
                "FROM CO_Documento a " +
                    "INNER JOIN CO_DocumentoDetalle b ON b.CompaniaSocio = a.CompaniaSocio AND b.TipoDocumento = a.TipoDocumento AND b.NumeroDocumento = a.NumeroDocumento " +
                    "INNER JOIN MA_FormadePago d ON a.FormadePago = d.FormadePago " +
                    "LEFT JOIN WH_ItemAlmacenLote f ON b.ItemCodigo = f.Item AND f.Condicion = 0 AND b.Lote = f.Lote AND b.ItemSerie = f.LoteFabricacion AND b.AlmacenCodigo = f.AlmacenCodigo " +
                "WHERE a.CompaniaSocio = '01000000' ";

            if (datos.TipoDocumento == "F")
                query = $"{query} AND a.TipoDocumento <> 'PE'";
            else
                query = $"{query} AND a.TipoDocumento = 'PE'";

            if (!string.IsNullOrEmpty(datos.NumeroDocumento))
                query = $"{query} AND RTRIM(a.NumeroDocumento) Like @NumeroDocumento ";

            if (!string.IsNullOrEmpty(datos.OrdenFabricacion))
                query = $"{query} AND b.Lote = @OrdenFabricacion";

            if (!string.IsNullOrEmpty(datos.Lote))
                query = $"{query} AND b.ItemSerie = @Lote";

            if (datos.IdCliente != 0)
                query = $"{query} AND a.ClienteNumero = @IdCliente ";

            if (datos.FechaInicio != null && datos.FechaFin != null)
                query = $"{query} AND CAST(a.FechaDocumento AS DATE) BETWEEN @FechaInicio AND @FechaFin";

            query = $"{query} ORDER BY a.TipoDocumento, a.NumeroDocumento, b.Linea ASC ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaProtocolo = await context.QueryAsync<DetalleProtocoloAnalisis>(query, datos);
            }

            return listaProtocolo.ToList();

        }

        public async Task<List<DetalleProtocoloAnalisis>> ListaProtocolosPorGuiaRemision(DatosProtocoloAnalisisListado datos)
        {
            IEnumerable<DetalleProtocoloAnalisis> listaProtocolo = new List<DetalleProtocoloAnalisis>();
            string query = "SELECT	a.GuiaNumero NumeroDocumento, a.FechaDocumento, null FechaVencimiento, " +
                "RTRIM(a.DestinatarioNombre) ClienteNombre, RTRIM(b.ItemCodigo) ItemCodigo, RTRIM(b.Descripcion) Descripcion, b.Cantidad CantidadPedida, " +
                "ISNULL(substring(e.referencianumero, 1, 8), b.Lote) Lote, e.FechaExpiracion, " +
                "RTRIM(b.Lote) AS OrdenFabricacion, RTRIM(a.Comentarios) Comentarios " +
                "FROM WH_GuiaRemision a " +
                "INNER JOIN WH_GuiaRemisionDetalle b ON b.CompaniaSocio = a.CompaniaSocio AND b.SerieNumero = a.SerieNumero AND b.GuiaNumero = a.GuiaNumero " +
                "LEFT JOIN EP_ProgramacionLote e ON b.CompaniaSocio = e.CompaniaSocio AND b.Lote = e.NumeroLote " +
                "WHERE a.CompaniaSocio = '01000000'";
            
            if (!string.IsNullOrEmpty(datos.NumeroDocumento))
                query = $"{query} AND RTRIM(a.GuiaNumero) LIKE @NumeroDocumento";

            if (!string.IsNullOrEmpty(datos.OrdenFabricacion))
                query = $"{query} AND b.Lote = @OrdenFabricacion";

            if (!string.IsNullOrEmpty(datos.Lote))
                query = $"{query} AND SUBSTRING( e.referencianumero, 1, 8) = @lote";

            if(datos.IdCliente != 0)
                query = $"{query} AND a.Destinatario = @IdCliente ";

            if (datos.FechaInicio != null && datos.FechaFin != null)
                query = $"{query} AND CAST(a.FechaDocumento AS DATE) BETWEEN @FechaInicio AND @FechaFin";

            query = $"{query} ORDER BY a.GuiaNumero, b.Secuencia ASC";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaProtocolo = await context.QueryAsync<DetalleProtocoloAnalisis>(query, datos);
            }

            return listaProtocolo.ToList();

        }

        public async Task<List<DetalleProtocoloAnalisis>> ListaProtocolosSinTipoDocumento(string ordenFabricacion, string lote)
        {
            IEnumerable<DetalleProtocoloAnalisis> listaProtocolo = new List<DetalleProtocoloAnalisis>();

            string query = "SELECT null  NumeroDocumento, null FechaDocumento, CAST('1900-01-01 00:00:00.000' AS DATE) FechaVencimiento, " +
                "null ClienteNombre, RTRIM(b.ITEM) ItemCodigo, RTRIM(b.descripcionlocal) Descripcion, 0 CantidadPedida, " +
                "SUBSTRING(a.referencianumero, 1, 8) AS Lote, ISNULL(a.FechaExpiracion, '1900-01-01 00:00:00.000') AS FechaExpiracion, " +
                "a.NUMEROLOTE AS OrdenFabricacion, null Comentarios " +
                "FROM ep_programacionlote a INNER JOIN WH_ItemMast b ON b.Item = a.Item WHERE a.CompaniaSocio = '01000000'";

            if (!string.IsNullOrEmpty(ordenFabricacion))
                query = $"{query} AND a.NumeroLote = @OrdenFabricacion";

            if (!string.IsNullOrEmpty(lote))
                query = $"{query} AND SUBSTRING( a.referencianumero, 1, 8) = @Lote ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaProtocolo = await context.QueryAsync<DetalleProtocoloAnalisis>(query, new { ordenFabricacion, lote });
            }

            return listaProtocolo.ToList();

        }

        public async Task<List<DetalleProtocoloAnalisis>> ListaProtocolosPorCotizacion(string numeroDocumento, int idCliente, DateTime? fechaInicio, DateTime? fechaFin)
        {
            IEnumerable<DetalleProtocoloAnalisis> listaProtocolo = new List<DetalleProtocoloAnalisis>();

            string query = "SELECT RTRIM(a.NumeroDocumento) NumeroDocumento, a.FechaDocumento, a.FechaVencimiento, " +
                "RTRIM(a.ClienteNombre) ClienteNombre, RTRIM(b.ItemCodigo) ItemCodigo, b.Descripcion, b.CantidadPedida,RTRIM(b.lote) Lote, " +
                "null AS FechaExpiracion, RTRIM(b.OrdenFabricacion) OrdenFabricacion, RTRIM(a.Comentarios) Comentarios " +
                "FROM CO_Cotizacion a " +
                "INNER JOIN( SELECT * FROM( " +
                    "SELECT b.FechaIngreso, a.CompaniaSocio, a.NumeroDocumento, a.ItemCodigo, a.Descripcion, a.UnidadCodigo, a.CantidadPedida, a.Monto, a.Linea, " +
                        "b.Lote OrdenFabricacion, SUBSTRING(LoteFabricacion, 1, 8) Lote, " +
                        "ROW_NUMBER() OVER(PARTITION BY a.ItemCodigo ORDER BY(LEFT(LoteFabricacion, 1) + RIGHT(RTRIM(LoteFabricacion), 1)) DESC) AS Row " +
                    "FROM CO_CotizacionDetalle a " +
                        "LEFT JOIN WH_ItemAlmacenLote b ON b.item = a.ItemCodigo AND lote<> '00' " +
                    "WHERE NumeroDocumento = @numeroDocumento AND LoteFabricacion IS NOT NULL AND Lote IS NOT NULL AND " +
                        "LEFT(LoteFabricacion, 2) NOT IN('OT', 'PE') AND AlmacenCodigo IN('ALMPRT', 'ALMVTO', 'ALMLICIT', 'ALMDRO', 'CONTRCALID') " +
                ") x WHERE Row = 1) b ON b.CompaniaSocio = a.CompaniaSocio AND b.NumeroDocumento = a.NumeroDocumento " +
                "WHERE a.CompaniaSocio = '01000000' AND a.NumeroDocumento = @numeroDocumento ";

            if (idCliente != 0)
                query = $"{query} AND a.ClienteNumero = @idCliente";

            if (fechaInicio != null && fechaFin != null)
                query = $"{query} AND CAST(a.FechaDocumento AS DATE) BETWEEN @FechaInicio AND @FechaFin";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                listaProtocolo = await context.QueryAsync<DetalleProtocoloAnalisis>(query, new { numeroDocumento, idCliente, fechaInicio, fechaFin });
            }

            return listaProtocolo.ToList();
        }

        public async Task<List<ValidacionProtocoloDTO>> ValidarSiTieneProtocolo_OF(string ordenesFabricacion)
        {
            IEnumerable<ValidacionProtocoloDTO> validacion = new List<ValidacionProtocoloDTO>();
            
            using (var context = new SqlConnection(_appConfig.contextSpring))
            {
                validacion = await context.QueryAsync<ValidacionProtocoloDTO>("usp_ValidarSiTieneProtocolo_OF", new { ordenesFabricacion }, commandType: CommandType.StoredProcedure);
            }

            return validacion.ToList();
        }

        public async Task<List<DetalleClientes>> ListarClientes()
        {
            List<DetalleClientes> result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_ListarClientes", commandType: CommandType.StoredProcedure))
                {
                    result = result_db.Read<DetalleClientes>().ToList();
                }
                connection.Dispose();
            }
            return result;
        }


        public async Task<IEnumerable<FormatoLicitaciones>> ListarDocumentoLicitacion(DatosFormatoDocumentoLicitacion dato)
        {
            IEnumerable<FormatoLicitaciones> result = new List<FormatoLicitaciones>();

            string script = "SELECT FechaDocumento,RTRIM(Destinatario) Destinatario,RTRIM(SerieNumero) SerieNumero, RTRIM(GuiaNumero) GuiaNumero, DestinatarioDireccion , DestinatarioDireccionSecuencia , " +
                            "RTRIM(FacturaNumero) FacturaNumero , RTRIM(AlmacenCodigo) AlmacenCodigo, RTRIM(ReferenciaNumeroPedido) ReferenciaNumeroPedido, RTRIM(Comentarios) Comentarios " +
                            " FROM WH_GuiaRemision WHERE FechaDocumento >= CONCAT(@fechainicio ,' 00:00:00.000') AND FechaDocumento<= CONCAT(@fechafinal ,' 23:59:00.000') AND  Estado<>'AN' AND Destinatario = @idcliente ";
            using (SqlConnection connection = new SqlConnection(_appConfig.contextSpring))
            {
                result = await connection.QueryAsync<FormatoLicitaciones>(script, new { dato.fechainicio, dato.fechafinal, dato.idcliente });

            }

            return result;
        }

        public async Task<FormatoReporteGuiaRemisionesModel> NumerodeGuiaLicitacion(string dato)
        {

            FormatoReporteGuiaRemisionesModel result = new FormatoReporteGuiaRemisionesModel();

            string script = "SELECT p.DescripcionComercial DescripcionProceso, p.DescripcionComercialDetalle, d.NombredelaUnidadEjecutora as Region, e.OrdenCompra, e.Pecosa, d.NumeroContrato Contrato, RTRIM(e.NumeroEntrega) NumeroEntrega, RTRIM(per.NombreCompleto) ClienteNombre, d.NombreRegion, CONCAT(RTRIM(g.SerieNumero), '-', RTRIM(g.GuiaNumero)) GuiaNumero , p.CantItems "
                         + " FROM TBMLIProceso p "
                         + " INNER JOIN TBDLIProcesoDetalle d on p.IdProceso = d.IdProceso "
                         + " INNER JOIN TBDLIProcesoEntrega E ON D.IdDetalle = E.IdDetalle "
                         + " INNER JOIN PROD_UNILENE2..WH_GuiaRemision g on e.NumeroEntrega = SUBSTRING(g.ReprogramacionPuntoPartida, 1, CHARINDEX('-', g.ReprogramacionPuntoPartida) - 1) "
                         + " AND e.OrdenCompra = g.ReferenciaNumeroOrden and e.Pecosa = SUBSTRING(g.ReprogramacionPuntoPartida, CHARINDEX('-', g.ReprogramacionPuntoPartida) + 1, LEN(g.ReprogramacionPuntoPartida) - CHARINDEX('-', g.ReprogramacionPuntoPartida)) "
                         + " INNER JOIN PROD_UNILENE2..WH_GuiaRemisionDetalle h ON g.GuiaNumero = h.GuiaNumero AND h.SerieNumero = g.SerieNumero"
                         + " INNER JOIN PROD_UNILENE2..PersonaMast per ON per.Persona = g.Destinatario "
                         + " WHERE CONCAT(RTRIM(g.SerieNumero) ,'-', g.GuiaNumero) IN(" + dato + ") "
                         + " SELECT D.NumeroItem, RTRIM(h.Descripcion) Descripcion, ISNULL(RTRIM(prm.Presentacion),'') CaractervaluesDescripcion, RTRIM(IM.UnidadCodigo) UnidadCodigo, "
                         + " CONVERT(INT,d.CantidadRequerida) CantidadRequerida, CONVERT(INT,e.cantidad) cantidad, CONVERT(INT,h.Cantidad) CantidadGRD, concat(RTRIM(g.SerieNumero), '-', RTRIM(g.GuiaNumero)) Guia, RTRIM(h.Lote) Lote, "
                         + " pl.Fechavencimiento FechaExpiracion, ISNULL(prm.RegistroSanitario,'') RegistroSanitario, ISNULL(prm.Temperatura,'') Temperatura, IIF(LEFT(h.ItemCodigo, 1)='P',h.Lote,'S/N') Protocolo, prm.NumeroMuestreo, prm.NumeroEnsayo "
                         + " FROM TBMLIProceso p "
                         + "INNER JOIN TBDLIProcesoDetalle d ON p.IdProceso = d.IdProceso "
                         + "INNER JOIN TBDLIProcesoEntrega E ON D.IdDetalle = E.IdDetalle "
                         + "INNER JOIN PROD_UNILENE2..WH_GuiaRemision g ON e.NumeroEntrega = SUBSTRING(g.ReprogramacionPuntoPartida, 1, CHARINDEX('-', g.ReprogramacionPuntoPartida) - 1) "
                         + "and e.OrdenCompra = g.ReferenciaNumeroOrden AND e.Pecosa = SUBSTRING(g.ReprogramacionPuntoPartida, CHARINDEX('-', g.ReprogramacionPuntoPartida) + 1, LEN(g.ReprogramacionPuntoPartida) - CHARINDEX('-', g.ReprogramacionPuntoPartida)) "
                         + "INNER JOIN PROD_UNILENE2..WH_GuiaRemisionDetalle h ON g.GuiaNumero = h.GuiaNumero and h.SerieNumero = g.SerieNumero "
                         + "INNER JOIN PROD_UNILENE2..WH_ItemMast im ON h.itemcodigo = im.item "
                         + "INNER JOIN PROD_UNILENE2..WH_CaracteristicaValues C ON IM.CaracteristicaValor01 = C.Valor AND C.Caracteristica = '01' "
                         + "INNER JOIN PROD_UNILENE2..WH_ItemAlmacenLote pl ON pl.Lote = h.Lote AND g.almacencodigo=pl.almacencodigo AND h.ItemCodigo=pl.Item "
                         + "INNER JOIN TBDLIProcesoProgramacionMuestras prm ON p.idproceso = prm.idproceso AND d.NumeroItem = prm.NumeroItem AND e.NumeroEntrega = prm.NumeroEntrega "
                         + "INNER JOIN PROD_UNILENE2..PersonaMast per ON per.Persona = g.Destinatario "
                         + "WHERE CONCAT(RTRIM(g.SerieNumero) ,'-', g.GuiaNumero) IN(" + dato + ") ";
            

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (SqlMapper.GridReader reporte = await connection.QueryMultipleAsync(script))
                {
                    result.CabeceraReporteGuiaRemision = reporte.Read<CReporteGuiaRemisionModel>().ToList();
                    result.DetalleReporteGuiaRemision = reporte.Read<DReportGuiaRemisionModel>().ToList();
                }
            }
            return result;
        }

        public async Task<IEnumerable<FormatoReporteProtocoloModel>> NumerodeGuiaProtocolo(string dato)
        {
            string nuevovalor = dato.Replace("'","");

            IEnumerable<FormatoReporteProtocoloModel> result = new List<FormatoReporteProtocoloModel>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result= await connection.QueryAsync<FormatoReporteProtocoloModel>("usp_Imprimir_Protocolo_Analisis_Lote", new { GUIAS=nuevovalor }, commandType: CommandType.StoredProcedure);
            }
            return result;
        }

        public async Task<DatoPedidoDocumentoModel> NumeroPedido(string pedido)
        {
            DatoPedidoDocumentoModel result  = new DatoPedidoDocumentoModel();
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                 result = await context.QueryFirstOrDefaultAsync<DatoPedidoDocumentoModel>("usp_ObtenerInformacionPedido", new { pedido }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task RegistrarRotuladosPedido(DatosEstructuraNumeroRotuloModel dato, int idUsuario)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync("usp_RegistarRotuladoPedido", new { dato.numeroDocumento, dato.Rexterno,dato.Rinterno, idUsuario }, commandType: CommandType.StoredProcedure);
            }   
        }

        public async Task<IEnumerable<FormatoGuiaPorFacturarModel>>  ListarGuiaporFacturar(DatosEstructuraGuiaPorFacturarModel dato)
        {
            IEnumerable<FormatoGuiaPorFacturarModel> lista = new List<FormatoGuiaPorFacturarModel>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {   
                lista=await context.QueryAsync<FormatoGuiaPorFacturarModel>("usp_listar_guias_por_facturar", new { dato.Territorio, dato.FechaInicio,dato.FechaFin,dato.destinatario,dato.Tipo }, commandType: CommandType.StoredProcedure);
            }

            return lista;
        }

        public async Task<IEnumerable<FormatoGuiaPorFacturarGeneralModel>> ListarGuiaporFacturarGeneral(DatosEstructuraGuiaPorFacturarModel dato)
        {
            IEnumerable<FormatoGuiaPorFacturarGeneralModel> lista = new List<FormatoGuiaPorFacturarGeneralModel>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await context.QueryAsync<FormatoGuiaPorFacturarGeneralModel>("usp_listar_guias_por_facturar_general", new { destinatario = dato.destinatario  }, commandType: CommandType.StoredProcedure);
            }

            return lista;
        }

        public async Task<(List<ProtocoloCabeceraModel> cabeceras, List<ProtocoloDetalleModel> detalles)> ObtenerDatosReporteProtocolo(string ordenFabricacion)
        {
            List<ProtocoloCabeceraModel> cabeceras = new List<ProtocoloCabeceraModel>();
            List<ProtocoloDetalleModel> detalles = new List<ProtocoloDetalleModel>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                using (var result_db = await context.QueryMultipleAsync("usp_DatosReporteAnalisisProtocolo", new { ordenFabricacion }, commandType: CommandType.StoredProcedure))
                {
                    cabeceras = result_db.Read<ProtocoloCabeceraModel>().ToList();
                    detalles = result_db.Read<ProtocoloDetalleModel>().ToList();
                }
            }

            return (cabeceras, detalles);
        }

        public async Task<RegistroRecepcionGuiaResponseDTO> RegistrarAdministracionGuia(string serie, string documento, int idUsuario, string tipoRegistro)
        {
            RegistroRecepcionGuiaResponseDTO datosGuia = new RegistroRecepcionGuiaResponseDTO();

            string query = null;

            if (tipoRegistro == "I")
            {
                query = "UPDATE WH_GuiaRemision SET ComentariosEntrega = '1', AgenciaTransporte=@idUsuario, FechaReprogramacion1 = GETDATE() WHERE CompaniaSocio = '01000000' AND SerieNumero=@serie AND GuiaNumero=@documento";
            }
            if (tipoRegistro == "O")
            {
                query = "UPDATE WH_GuiaRemision SET FechaReprogramacion2 = GETDATE() WHERE CompaniaSocio = '01000000' AND SerieNumero=@serie AND GuiaNumero=@documento";
            }

            string consulta = "SELECT RTRIM(SerieNumero) + '-' + RTRIM(GuiaNumero) GuiaNumero, DestinatarioNombre Destinatario, FacturaNumero Factura, FechaDocumento FROM WH_GuiaRemision WHERE CompaniaSocio = '01000000' AND SerieNumero=@serie AND GuiaNumero=@documento";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                await context.ExecuteAsync(query, new { serie, documento, idUsuario });
                datosGuia = await context.QueryFirstOrDefaultAsync<RegistroRecepcionGuiaResponseDTO>(consulta, new { serie, documento });
            }

            return datosGuia;
        }

        public async Task<(string guiaNumero, string comentariosEntrega, DateTime? FechaEnvioAlmacen)> ObtenerGuiaRegistrada (string serie, string documento)
        {
            (string guiaNumero, string comentariosEntrega, DateTime? FechaEnvioAlmacen) response;
            string query = "SELECT GuiaNumero, RTRIM(ComentariosEntrega) EstadoEntrega, FechaReprogramacion2 FechaEnvioAlmacen FROM WH_GuiaRemision WHERE CompaniaSocio = '01000000' AND SerieNumero=@serie AND GuiaNumero=@documento";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                response = await context.QueryFirstOrDefaultAsync<(string guiaNumero, string comentariosEntrega, DateTime? FechaEnvioAlmacen)>(query, new { serie, documento });
            }

            return response;
        }

    }
}
