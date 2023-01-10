using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class LicitacionesRepository : ILicitacionesRepository
    {
        private readonly IAppConfig _appConfig;
        public LicitacionesRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido)
        {

            IEnumerable<ListarDetallePedido> result = new List<ListarDetallePedido>();
            string script = "SELECT RTRIM(a.NumeroDocumento) Pedido, a.Linea, RTRIM(a.ItemCodigo) Item, RTRIM(a.Descripcion) Descripcion, a.CantidadPedida , a.PrecioUnitario, a.Monto " +
                            "FROM CO_DocumentoDetalle a INNER JOIN CO_Documento b ON  a.CompaniaSocio=b.CompaniaSocio AND a.TipoDocumento=b.TipoDocumento AND a.NumeroDocumento=b.NumeroDocumento " +
                            "WHERE  a.TipoDocumento = 'PE' AND  a.NumeroDocumento = @Pedido  AND b.ClienteNumero=2317";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<ListarDetallePedido>(script, new { Pedido });
            }
            return result;
        }

        public async Task<int> RegistrarProceso(DatoFormatoProcesoModel dato)
        {
            int result;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql1 = "INSERT INTO TBMLIProceso(DescripcionProceso, Cliente)" +
                             "VALUES(@Proceso, @Cliente)";

                result  = await context.ExecuteAsync(sql1, new { dato.Proceso, dato.Cliente });
                        

            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoDistribuccionLP>> ListarDistribuccionProceso(int NumeroProceso, string Item , string Mes)
        {
            IEnumerable<DatosFormatoDistribuccionLP> result = new List<DatosFormatoDistribuccionLP>();

            string sql1 = "SELECT e.IdEntrega, p.DescripcionProceso, D.NombreDiresa, CodigoAlmacen, D.PuntosdeEntrega, D.Tipodeusuario, D.NumeroItem, D.CodItem, D.CodigoSISMED, D.DescripcionItem, D.PrecioUnitario, D.CantidadRequerida, e.Cantidad, e.OrdenCompra, e.Pecosa " +
                           "FROM TBMLIProceso P " +
                           "INNER JOIN TBDLIProcesoDetalle D ON P.IdProceso = D.IdProceso " +
                           "INNER JOIN TBDLIProcesoEntrega E ON D.IdDetalle = E.IdDetalle " +
                           "WHERE E.NumeroEntrega = @Mes AND P.IdProceso = @NumeroProceso  AND D.CodItem= IIF(@Item = '' OR @Item is null, D.CodItem, @Item)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoDistribuccionLP>(sql1, new { Mes, NumeroProceso, Item });
               
            }

            return result;
        }

        public async Task<IEnumerable<string>> ObtenerTipoUsuario(int NumeroProceso, int Item, string Mes)
        {
            IEnumerable<string> result = new List<string>();

            string sql1 =  "SELECT D.Tipodeusuario " +
                           "FROM TBMLIProceso P " +
                           "INNER JOIN TBDLIProcesoDetalle D ON P.IdProceso = D.IdProceso " +
                           "INNER JOIN TBDLIProcesoEntrega E ON D.IdDetalle = E.IdDetalle " +
                           "WHERE E.NumeroEntrega = @Mes AND P.IdProceso = @NumeroProceso  AND D.NumeroItem=@Item  " +
                           "GROUP BY D.Tipodeusuario";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<string>(sql1, new { Mes, NumeroProceso, Item });

            }

            return result;
        }


        public async Task<DatosFormatoBuscarOrdenCompraLicitacionesModel> BuscarOrdenCompraLicitaciones(int NumeroProceso, int NumeroEntrega, int Item, string TipoUsuario)
        {
            DatosFormatoBuscarOrdenCompraLicitacionesModel result = new DatosFormatoBuscarOrdenCompraLicitacionesModel() ;

            string sql1 = "SELECT TOP 1 NumeroOrden numeroOC, CantidadOrden cantidadOC FROM TBDLIOrdenCompra WHERE IdProceso= @NumeroProceso AND NumeroEntrega=@NumeroEntrega AND NumeroItem=@Item AND TipoUsuario=@TipoUsuario AND ESTADO=1 ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryFirstOrDefaultAsync<DatosFormatoBuscarOrdenCompraLicitacionesModel>(sql1,new { NumeroProceso, NumeroEntrega , Item, TipoUsuario });
            }
            return result;
        }


        public async Task<int> RegistrarOrdenCompra(DatoFormatoRegistrarOrdenCompraLicitaciones dato ,int idUsuario)
        {
            int result;
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.ExecuteAsync("usp_Licitaciones_Registrar_OrdenCompra", new { dato.idProceso,dato.NumeroEntrega,dato.NumeroOC,dato.Usuario,dato.CantidadOC,dato.Item, idUsuario },commandType: CommandType.StoredProcedure);
            }
            return result;
        }


        public async Task RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> dato)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "UPDATE TBDLIProcesoEntrega SET OrdenCompra=@OrdenCompra , Pecosa=@Pecosa WHERE IdEntrega=@idEntrega";
                await context.ExecuteAsync(sql, dato);
            }   
        }


        public async Task<IEnumerable<ListarProcesoEntity>> ListarProceso(int idClient)
        {
            IEnumerable<ListarProcesoEntity> result = new List<ListarProcesoEntity>();

            string query = "SELECT IdProceso,DescripcionProceso,DescripcionComercial,DescripcionComercialDetalle," +
                "CantItems, Cliente FROM TBMLIProceso";

            if (idClient != 0)
                query = $"{query} WHERE Cliente = @idClient";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<ListarProcesoEntity>(query, new { idClient  });
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoProgramacionMuestraModel>> ListarProgramaMuestraLIP(int IdProceso,string NumeroEntrega)
        {
            IEnumerable<DatosFormatoProgramacionMuestraModel> result = new List<DatosFormatoProgramacionMuestraModel>();

            string sql1 = "SELECT M.IdProgramacion,M.NumeroEntrega,M.NumeroItem,RTRIM(D.DescripcionItem) DescripcionItem,M.CodItem,M.NumeroMuestreo,M.Presentacion,M.Temperatura,M.RegistroSanitario,M.NumeroEnsayo, M.Protocolo, P.IdProceso FROM TBMLIProceso P " +
                            "INNER JOIN TBDLIProcesoDetalle D ON P.IdProceso = D.IdProceso "+
                            "INNER JOIN TBDLIProcesoProgramacionMuestras M ON P.IdProceso = M.IdProceso AND D.NumeroItem = M.NumeroItem "+
                            "WHERE P.IdProceso = @IdProceso AND M.NumeroEntrega = @NumeroEntrega " +
                            "GROUP BY M.IdProgramacion,M.NumeroEntrega,M.NumeroItem,M.CodItem,D.DescripcionItem,M.NumeroMuestreo,M.Presentacion,M.Temperatura,M.RegistroSanitario,M.NumeroEnsayo, M.Protocolo,M.IdProgramacion,p.IdProceso " +
                            "ORDER BY M.NumeroItem";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoProgramacionMuestraModel>(sql1, new { IdProceso , NumeroEntrega });
            }
            return result;
        }



        public async Task RegistrarProgramacionMuestreo(List<DatosFormatoMuestraEnsayoLIP> dato)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "UPDATE TBDLIProcesoProgramacionMuestras SET NumeroMuestreo=@numeroMuestreo , NumeroEnsayo=@numeroEnsayo  WHERE IdProgramacion=@idProgramacion " +
                             "UPDATE TBDLIProcesoProgramacionMuestras SET Protocolo=@protocolo , Temperatura=@temperatura , RegistroSanitario=@registroSanitario, Presentacion=@presentacion  WHERE NumeroItem=@numeroItem AND IdProceso=@idProceso";
                await context.ExecuteAsync(sql, dato);
            }

        }



        public async Task<IEnumerable<ListarGuiaInformeLPModel>> ListarGuiaInformacion(string NumeroEntrega,string OrdenCompra)
        {
            IEnumerable<ListarGuiaInformeLPModel> result = new List<ListarGuiaInformeLPModel>();

            string sql1 = "SELECT RTRIM(g.SerieNumero) SerieNumero , RTRIM(g.GuiaNumero) GuiaNumero,g.fechadocumento,RTRIM(g.ReferenciaNumeroOrden) OrdenCompra,g.estado Estado,RTRIM(g.COMENTARIOS) Comentario,g.ReprogramacionPuntoPartida AS 'EntregaPecosa',g.ReferenciaNumeroOrden , " +
                "IIF(o.FECHA_RETORNO IS NULL, 'SIN ORDEN DE SERVICIO', IIF(O.FECHA_RETORNO = '1900-01-01 00:00:00.000', 'SIN RETORNO', 'OK')) 'EstadoLogistica',RTRIM(o.comentarios) ComentarioSalida " +
                "FROM PROD_UNILENE2..WH_GuiaRemision g "+
                "FULL OUTER JOIN UNILENE_REPORTEADOR..TLOG_PLAN_ORDEN_SERVICIO_D O " +
                "ON g.SerieNumero = o.serie and g.GuiaNumero = o.numero_documento " +
                "INNER JOIN PROD_UNILENE2..wh_guiaremisiondetalle D on g.serienumero = d.serienumero and g.guianumero = d.guianumero "+
                "INNER JOIN PROD_UNILENE2..WH_TransaccionHeader t on d.referenciatipodocumento = t.TipoDocumento and d.ReferenciaNumeroDocumento = t.NumeroDocumento "+
                "WHERE g.Destinatario = 2317 and t.TransaccionCodigo <> 'SDE' and RTRIM(g.ReferenciaNumeroOrden) = RTRIM(@OrdenCompra) and SUBSTRING(g.ReprogramacionPuntoPartida,1,1) = @NumeroEntrega ";


            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<ListarGuiaInformeLPModel>(sql1, new{ OrdenCompra, NumeroEntrega });
            }
            return result;
        }


        public async Task<IEnumerable<EstructuraListaContratoProceso>> ListarContratoProceso(string proceso)
        {
            IEnumerable<EstructuraListaContratoProceso> result = new List<EstructuraListaContratoProceso>();
            string sql1 = "SELECT idproceso, tipodeusuario, numeroitem, descripcionitem, ISNULL(NumeroContrato,'') NumeroContrato FROM TBDLIProcesoDetalle " +
                           "WHERE idproceso=@proceso GROUP BY idproceso,tipodeusuario,numeroitem,descripcionitem,NumeroContrato";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<EstructuraListaContratoProceso>(sql1, new { proceso});
            }
            return result;
        }


        public async Task RegistrarContratoProceso(List<DatosRequestFormatoContratoProcesoModel> dato)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "UPDATE TBDLIProcesoDetalle SET NumeroContrato = @numeroContrato  WHERE IdProceso=@idproceso AND Tipodeusuario=@tipodeusuario AND NumeroItem=@numeroitem";
                await context.ExecuteAsync(sql, dato);
            }

        }

        public async Task<IEnumerable<DatosFormatodashboardLicitaciones>> DashboardLicitacionesExportar()
        {
            IEnumerable<DatosFormatodashboardLicitaciones> result = new List<DatosFormatodashboardLicitaciones>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result= await context.QueryAsync<DatosFormatodashboardLicitaciones>("usp_Listar_Informacion_DashboardLicitaciones", commandType: CommandType.StoredProcedure);
                
            }

            return result;
        }

    }
}
