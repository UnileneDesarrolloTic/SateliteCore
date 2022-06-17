﻿using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
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

        public async Task RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> dato)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "UPDATE TBDLIProcesoEntrega SET OrdenCompra=@OrdenCompra , Pecosa=@Pecosa WHERE IdEntrega=@idEntrega";
                await context.ExecuteAsync(sql, dato);
            }   
        }


        public async Task<IEnumerable<ListarProcesoEntity>> ListarProceso()
        {
            IEnumerable<ListarProcesoEntity> result = new List<ListarProcesoEntity>();

            string sql1 = "SELECT IdProceso, DescripcionProceso ,DescripcionComercial,DescripcionComercialDetalle ,CantItems   FROM TBMLIProceso";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<ListarProcesoEntity>(sql1);
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoProgramacionMuestraModel>> ListarProgramaMuestraLIP(int IdProceso,string NumeroEntrega)
        {
            IEnumerable<DatosFormatoProgramacionMuestraModel> result = new List<DatosFormatoProgramacionMuestraModel>();

            string sql1 = "SELECT M.IdProgramacion,M.NumeroEntrega,M.NumeroItem,D.DescripcionItem,M.CodItem,M.NumeroMuestreo,M.NumeroEnsayo,P.IdProceso FROM TBMLIProceso P " +
                            "INNER JOIN TBDLIProcesoDetalle D ON P.IdProceso = D.IdProceso "+
                            "INNER JOIN TBDLIProcesoProgramacionMuestras M ON P.IdProceso = M.IdProceso AND D.NumeroItem = M.NumeroItem "+
                            "WHERE P.IdProceso = @IdProceso AND M.NumeroEntrega = @NumeroEntrega " +
                            "GROUP BY M.IdProgramacion,M.NumeroEntrega,M.NumeroItem,M.CodItem,D.DescripcionItem,M.NumeroMuestreo,M.NumeroEnsayo,M.IdProgramacion,p.IdProceso " +
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
                string sql = "UPDATE TBDLIProcesoProgramacionMuestras SET NumeroMuestreo=@numeroMuestreo , NumeroEnsayo=@numeroEnsayo WHERE IdProgramacion=@idProgramacion";
                await context.ExecuteAsync(sql, dato);
            }

        }
    }
}
