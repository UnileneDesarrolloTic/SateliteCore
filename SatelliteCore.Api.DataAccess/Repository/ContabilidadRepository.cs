using System;
using MongoDB.Bson;
using MongoDB.Driver;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.IO;
using System.Text;
using SatelliteCore.Api.Models.Response.Contabilidad;
using SatelliteCore.Api.Models.Request.Contabildad;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ContabilidadRepository : IContabilidadRepository
    {
        private readonly IAppConfig _appConfig;

        public ContabilidadRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }


        public async Task<IEnumerable<DetraccionesEntity>> ListarDetraccion()
        {
            IEnumerable<DetraccionesEntity> result_db = new List<DetraccionesEntity>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result_db = await connection.QueryAsync<DetraccionesEntity>("usp_ListarDetraccionContabilidad", commandType: CommandType.StoredProcedure);
                connection.Dispose();
            }

            return result_db;
        }

        public int ProcesarDetraccionContabilidad(List<FormatoComprobantePagoDetraccion> dato)
        {
                
                using (var satelliteContext = new SqlConnection(_appConfig.contextSatelliteDB))
                {
                    foreach (var item in dato)
                    {
                        DynamicParameters parameter = new DynamicParameters();
                        parameter.Add("@TipoCuenta", item.TipoCuenta, DbType.String, ParameterDirection.Input);
                        parameter.Add("@NumeroCuenta", item.NumeroCuenta, DbType.String, ParameterDirection.Input);
                        parameter.Add("@NumeroConstancia", item.NumeroConstancia, DbType.String, ParameterDirection.Input);
                        parameter.Add("@PeriodoTributario", item.PeriodoTributario, DbType.String, ParameterDirection.Input);
                        parameter.Add("@RUC_Proveedor", item.RucProveedor, DbType.String, ParameterDirection.Input);
                        parameter.Add("@Nombre_Proveedor", item.NombreProveedor, DbType.String, ParameterDirection.Input);
                        parameter.Add("@TipoDocumentoAdquiriente", item.TipoDocumento, DbType.String, ParameterDirection.Input);
                        parameter.Add("@NumeroDocumentoAdquiriente", item.DocumentoAdquiriente, DbType.String, ParameterDirection.Input);
                        parameter.Add("@RazonSocialAdquiriente", item.RazonSocial, DbType.String, ParameterDirection.Input);
                        parameter.Add("@FechaPago", item.FechaPago, DbType.DateTime, ParameterDirection.Input);
                        parameter.Add("@MontoDeposito", item.MontoDeposito, DbType.Decimal, ParameterDirection.Input);
                        parameter.Add("@TipoBien", item.TipoBien, DbType.String, ParameterDirection.Input);
                        parameter.Add("@TipoOperacion", item.TipoOperacion, DbType.String, ParameterDirection.Input);
                        parameter.Add("@TipoComprobante", item.TipodeComprobante, DbType.String, ParameterDirection.Input);
                        parameter.Add("@SerieComprobante", item.Serie, DbType.String, ParameterDirection.Input);
                        parameter.Add("@NumeroComprobante", item.Numero, DbType.String, ParameterDirection.Input);
                        parameter.Add("@NumeroPagoDetracciones", item.PagoDetraccion, DbType.String, ParameterDirection.Input);
                     
                        satelliteContext.Execute("usp_RegistrarDetraccionesContabilidad", parameter, commandType: CommandType.StoredProcedure);


                     }

                        satelliteContext.Dispose();
                }
                
                return 1;

        }


        public async Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ConsultarProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato)
        {
            IEnumerable<DatosFormatoDatosProductoCostobase> result_db = new List<DatosFormatoDatosProductoCostobase>();
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result_db = await connection.QueryAsync<DatosFormatoDatosProductoCostobase>("usp_BaseCosto_ItemProducto_Analisis_costo", new{ dato.CodProducto, dato.NumeroCotizacion, dato.Opcion , dato.base64, dato.BusquedaExcel, dato.idfamilia,dato.idSubFamilia } , commandType: CommandType.StoredProcedure);
                connection.Dispose();
            }

            return result_db;
        }


        public async Task<IEnumerable<DatosFormatoRecetaItemComponente>> ConsultarRecetaProducto(string Item)
        {
            IEnumerable<DatosFormatoRecetaItemComponente> result_db = new List<DatosFormatoRecetaItemComponente>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result_db = await connection.QueryAsync<DatosFormatoRecetaItemComponente>("usp_Info_Receta_itemcomponente_MP",new { Item }, commandType: CommandType.StoredProcedure);
                connection.Dispose();
            }

            return result_db;
        }

        public async Task<IEnumerable<DatosFormatoComponentePrecioUnitario>> ListarItemComponentePrecio(DatosFormatosComponentPrecio dato)
        {
            IEnumerable<DatosFormatoComponentePrecioUnitario> result_db = new List<DatosFormatoComponentePrecioUnitario>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result_db = await connection.QueryAsync<DatosFormatoComponentePrecioUnitario>("usp_Listar_Item_Componente_CostoBase", new { dato.Linea,dato.Familia,dato.SubFamilia }, commandType: CommandType.StoredProcedure);
                connection.Dispose();
            }

            return result_db;
        }

        public async Task<(List<FormatoListadoInformacionTransaccionKardex>, int)> InformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato)
        {
            (List<FormatoListadoInformacionTransaccionKardex> Listado, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_InformacionTransaccionKardex", new {dato.Periodo,dato.Tipo,dato.Pagina,dato.RegistrosPorPagina }, commandType: CommandType.StoredProcedure))
                {
                    result.Listado = result_db.Read<FormatoListadoInformacionTransaccionKardex>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }
                connection.Dispose();
            }

            return result;
        }


        public async Task<string> RegistrarInformacionTransaccionKardex(DatoFormatoRegistrarTransaccionKardex docRegistrado)
        {
            IMongoDatabase integrationContext = new MongoClient(_appConfig.ContextMongoDB).GetDatabase("UnileneReporte");
            IMongoCollection<DatoFormatoRegistrarTransaccionKardex> _requestModel = integrationContext.GetCollection<DatoFormatoRegistrarTransaccionKardex>("TransaccionHistorica");
            await _requestModel.InsertOneAsync(docRegistrado);

            string idMongoDB = ""; //Transaccionkardex.GetValue("_id", "").ToString();
            return idMongoDB;
        }

        public async Task GuardarInformacionTransaccionKardex(string idMongoDB, string Tipo, string Periodo, bool CheckCierre, string usuario)
        {
            string sql = "INSERT INTO TBMCierreHistoricoContable (Codigo,Periodo,CheckCierre,CantidadTotal,MontoTotal,Tipo,Usuario,FechaRegistro) VALUES (@idMongoDB,@Tipo,@Periodo,@CheckCierre,0,0,@usuario,GETDATE())";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync(sql, new { idMongoDB, Tipo, Periodo, CheckCierre, usuario });
                connection.Dispose();
            }

            
        }


    }
}
