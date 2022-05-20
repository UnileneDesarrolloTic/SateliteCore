using System;
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

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ContabilidadRepository : IContabilidadRepository
    {
        private readonly IAppConfig _appConfig;

        public ContabilidadRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }


        public async Task<List<DetraccionesEntity>> ListarDetraccion()
        {
            List<DetraccionesEntity> ListaUsuarios = new List<DetraccionesEntity>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_ListarDetraccionContabilidad" , commandType: CommandType.StoredProcedure))
                {
                    ListaUsuarios = result_db.Read<DetraccionesEntity>().ToList();
                }

                connection.Dispose();
            }

            return ListaUsuarios;
        }

        public async Task<int> ProcesarDetraccionContabilidad(List<FormatoComprobantePagoDetraccion> dato)
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
                        using (var result_db = await satelliteContext.QueryMultipleAsync("usp_RegistrarDetraccionesContabilidad", parameter , commandType: CommandType.StoredProcedure))
                        {
                         
                        }
                    }
                    satelliteContext.Dispose();
                }
                
                return 1;

        }

      


    }
}
