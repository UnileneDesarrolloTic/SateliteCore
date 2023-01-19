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

        public async Task<(List<FormatoListadoInformacionTransaccionKardex>, FormatoCabeceraTransaccionKardex , int)> InformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato)
        {
            (List<FormatoListadoInformacionTransaccionKardex> detalle, FormatoCabeceraTransaccionKardex cabecera, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_InformacionTransaccionKardex", new { dato.Periodo, dato.Tipo, dato.Pagina, dato.RegistrosPorPagina }, commandType: CommandType.StoredProcedure))
                {
                    result.detalle = result_db.Read<FormatoListadoInformacionTransaccionKardex>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                    result.cabecera = result_db.Read<FormatoCabeceraTransaccionKardex>().FirstOrDefault();
                }
                connection.Dispose();
            }

            return result;
        }

        public async Task<bool> GuardarInformacionTransaccionKardex(DatoFormatoRegistrarTransaccionKardex docRegistrado, string usuario)
        {
            int idCabecera = 0;
            bool result = true;
            string sql = "SELECT IIF (EXISTS(SELECT 1 FROM TBMCierreHistorico WHERE Periodo=@Periodo AND Tipo=@Tipo AND CheckCierre=@CheckCierre AND Estado='A'), CAST(1 AS BIT), CAST(0 AS BIT)) AS exite";
            string sqlCabecera = "INSERT INTO TBMCierreHistorico (Periodo,CheckCierre,CantidadTotal,MontoTotal,Tipo,Estado,UsuarioCreacion,FechaRegistro) VALUES (@Periodo,@CheckCierre,@CCantidadTotal,@CMontoTotal,@Tipo,'A',@usuario,GETDATE()) " +
                                 " SELECT @@IDENTITY";


            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryFirstOrDefaultAsync<bool>(sql, new { docRegistrado.Periodo, docRegistrado.Tipo,docRegistrado.CheckCierre});

                if (!result)
                {
                    idCabecera = await connection.QueryFirstOrDefaultAsync<int>(sqlCabecera, new { docRegistrado.Periodo, docRegistrado.CheckCierre, docRegistrado.CCantidadTotal, docRegistrado.CMontoTotal, docRegistrado.Tipo, usuario });
                    foreach (FormatoListadoInformacionTransaccionKardex valor in docRegistrado.InformacionDetalle)
                        await connection.ExecuteAsync("usp_RegistrarDetalleReporteHistorico", new { docRegistrado.Tipo, idCabecera, docRegistrado.Periodo, valor.TipoDocumento, valor.NumeroDocumento, valor.TransaccionCodigo, valor.ReferenciaTipoDocumento, valor.ReferenciaNumeroDocumento, valor.Secuencia, valor.Item, valor.Lote, valor.Cantidad, valor.PrecioUnitario, valor.MontoTotal, valor.MontoTotalReal, valor.CodigoUnico }, commandType: CommandType.StoredProcedure);
                }

                connection.Dispose();
            }
            return result;
        }

        public async Task<IEnumerable<FormatoDatosCierreHistorico>> ListarInformacionReporteCierre(string Periodo)
        {
            IEnumerable<FormatoDatosCierreHistorico> Result =  new  List<FormatoDatosCierreHistorico>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                Result = await connection.QueryAsync<FormatoDatosCierreHistorico>("usp_Listado_ReporteCierre", new { Periodo }, commandType: CommandType.StoredProcedure);
                connection.Dispose();
            }

            return Result;
        }

        public async Task<IEnumerable<FormatoListadoInformacionTransaccionKardex>> ListarDetalleReporteCierre(int Id, string Periodo, string Tipo)
        {
            IEnumerable<FormatoListadoInformacionTransaccionKardex> Result = new List<FormatoListadoInformacionTransaccionKardex>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                Result = await connection.QueryAsync<FormatoListadoInformacionTransaccionKardex>("usp_ResultadoComparacionReporteCierre", new { Id, Periodo, Tipo }, commandType: CommandType.StoredProcedure);
                connection.Dispose();
            }

            return Result;
        }

        public async Task AnularReporteCierre(int Id, string usuario)
        {
            string sql = "UPDATE TBMCierreHistorico SET Estado = 'I', UsuarioModificacion = @usuario, FechaModificacion = GETDATE() WHERE Id = @Id";
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync(sql, new { Id, usuario });
                connection.Dispose();
            }
        }


    }
}
