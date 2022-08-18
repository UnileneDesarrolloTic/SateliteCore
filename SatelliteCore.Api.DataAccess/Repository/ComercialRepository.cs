using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
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

        public async Task<(List<DetalleProtocoloAnalisis>, int)> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos)
        {
            (List<DetalleProtocoloAnalisis> ListaProtocoloAnalisis, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_ListarProtocoloAnalisis", datos, commandType: CommandType.StoredProcedure))
                {
                    result.ListaProtocoloAnalisis = result_db.Read<DetalleProtocoloAnalisis>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }
                connection.Dispose();
            }
            return result;
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
                         + "INNER JOIN PROD_UNILENE2..WH_ItemAlmacenLote pl ON pl.lotefabricacion = h.Lote AND g.almacencodigo=pl.almacencodigo AND h.ItemCodigo=pl.Item AND SUBSTRING(pl.lote,1,2)<>'OT'"
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



        public async Task RegistrarGuiaporFacturar(DatoFormatoEstructuraGuiaFacturada dato, int idUsuario)
        {

            string script ="";

            if (dato.comentariosEntrega)
               script = "UPDATE PROD_UNILENE2..WH_GuiaRemision SET ComentariosEntrega='1' , AgenciaTransporte=@idUsuario , FechaReprogramacion1=GETDATE()  WHERE Destinatario=@destinatario AND SerieNumero=@serieNumero AND GuiaNumero=@guiaNumero";
            else
               script = "UPDATE PROD_UNILENE2..WH_GuiaRemision SET ComentariosEntrega='0', AgenciaTransporte=@idUsuario , FechaReprogramacion1=GETDATE() WHERE Destinatario=@destinatario AND SerieNumero=@serieNumero AND GuiaNumero=@guiaNumero";
                  
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { dato.destinatario, dato.serieNumero, dato.guiaNumero, dato.comentariosEntrega, idUsuario });
            }
        }



    }
}
