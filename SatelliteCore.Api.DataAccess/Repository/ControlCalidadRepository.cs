using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ControlCalidadRepository : IControlCalidadRepository
    {
        private readonly IAppConfig _appConfig;

        public ControlCalidadRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }
        public async Task<(List<CertificadoEsterilizacionEntity>, int)> ListarCertificados(DatosListarCertificadoPaginado datos)
        {
            (List<CertificadoEsterilizacionEntity> ListaCertificados, int totalRegistros) result;
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_ListaCertificadosEsterilizacion", datos, commandType: CommandType.StoredProcedure))
                {
                    result.ListaCertificados = result_db.Read<CertificadoEsterilizacionEntity>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }

                connection.Dispose();
            }

            return result;
        }

        public async Task<(List<LoteEntity>, int)> ListarLotes(DatosLote datos)
        {
            (List<LoteEntity> ListaLotes, int totalRegistros) result;
            datos.Descripcion = string.IsNullOrEmpty(datos.Descripcion) == true ? "" : datos.Descripcion;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string script = "SELECT Id, Descripcion, Lote, Expira FROM TBMCCLote WHERE Grupo = IIF(@Identificador = 0, Grupo, @Identificador) AND Descripcion LIKE '%' + @descripcion + '%' " +
                                " ORDER BY Id OFFSET (@pagina - 1) * @registrosPorPagina ROWS FETCH NEXT @registrosPorPagina ROWS ONLY; " +
                                "SELECT Count(1) FROM TBMCCLote WHERE Descripcion LIKE '%' + @descripcion + '%';";
                using (var result_db = await connection.QueryMultipleAsync(script, new { datos.Descripcion, datos.Identificador, datos.Pagina, datos.RegistrosPorPagina, }))
                {
                    result.ListaLotes = result_db.Read<LoteEntity>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }
                connection.Dispose();
            }

            return result;
        }

        public bool RegistrarCertificado(CertificadoEsterilizacionEntity certificado)
        {
            var si = true;
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                DynamicParameters parameter = new DynamicParameters();
                parameter.Add("@fechaEmision", certificado.FechaEmision, DbType.DateTime, ParameterDirection.Input);
                parameter.Add("@ordenServicio", certificado.OrdenServicio, DbType.String, ParameterDirection.Input);
                parameter.Add("@cliente", certificado.Cliente, DbType.String, ParameterDirection.Input);
                parameter.Add("@lote", certificado.Lote, DbType.String, ParameterDirection.Input);
                parameter.Add("@producto", certificado.Producto, DbType.String, ParameterDirection.Input);
                parameter.Add("@marca", certificado.Marca, DbType.String, ParameterDirection.Input);
                parameter.Add("@cantidad", certificado.Cantidad, DbType.Decimal, ParameterDirection.Input);
                parameter.Add("@equipo", certificado.Equipo, DbType.String, ParameterDirection.Input);
                parameter.Add("@cantidadUnidadMedida", certificado.CantidadUnidadMedida, DbType.String, ParameterDirection.Input);
                parameter.Add("@estado", certificado.Estado, DbType.String, ParameterDirection.Input);
                parameter.Add("@fechaInicio", certificado.FechaInicio, DbType.DateTime, ParameterDirection.Input);
                parameter.Add("@fechaTermino", certificado.FechaTermino, DbType.DateTime, ParameterDirection.Input);
                parameter.Add("@metodo", certificado.Metodo, DbType.String, ParameterDirection.Input);
                parameter.Add("@temperatura", certificado.Temperatura, DbType.String, ParameterDirection.Input);
                parameter.Add("@tiempoAireacion", certificado.TiempoAireacion, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@tiempoAireacionUnidadMedida", certificado.TiempoAireacionUnidadMedida, DbType.String, ParameterDirection.Input);
                parameter.Add("@tiempoExposicion", certificado.TiempoExposicion, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@tiempoExposicionUnidadMedida", certificado.TiempoExposicionUnidadMedida, DbType.String, ParameterDirection.Input);
                parameter.Add("@hrProceso", certificado.HRProceso, DbType.Decimal, ParameterDirection.Input);
                parameter.Add("@observaciones", certificado.Observaciones, DbType.String, ParameterDirection.Input);
                parameter.Add("@conclusion", certificado.Conclusion, DbType.String, ParameterDirection.Input);
                parameter.Add("@usuario", certificado.Usuario, DbType.String, ParameterDirection.Input);
                parameter.Add("@IDLoteCintaTestigo", certificado.IDLoteCintaTestigo, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@IDLoteIndicadorQuimico", certificado.IDLoteIndicadorQuimico, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@conformeCintaTestigo", certificado.ConformeCintaTestigo, DbType.Boolean, ParameterDirection.Input);
                parameter.Add("@conformeIndicadorQuimico", certificado.ConformeIndicadorQuimico, DbType.Boolean, ParameterDirection.Input);
                parameter.Add("@modeloTrazasOE", certificado.ModeloTrazasOE, DbType.String, ParameterDirection.Input);
                parameter.Add("@codigoTrazasOE", certificado.CodigoTrazasOE, DbType.String, ParameterDirection.Input);
                parameter.Add("@conformeTrazasOE", certificado.ConformeTrazasOE, DbType.Boolean, ParameterDirection.Input);
                parameter.Add("@tipoIB", certificado.TipoIB, DbType.String, ParameterDirection.Input);
                parameter.Add("@codigoIB", certificado.CodigoIB, DbType.String, ParameterDirection.Input);
                parameter.Add("@IDLoteIB", certificado.IDLoteIB, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@IBExpuestos", certificado.IBExpuestos, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@IBExpuestosResultado", certificado.IBExpuestosResultado, DbType.Boolean, ParameterDirection.Input);
                parameter.Add("@IBNoExpuestos", certificado.IBNoExpuestos, DbType.Int32, ParameterDirection.Input);
                parameter.Add("@IBNoExpuestosResultado", certificado.IBNoExpuestosResultado, DbType.Boolean, ParameterDirection.Input);
                parameter.Add("@conformeIB", certificado.ConformeIB, DbType.Boolean, ParameterDirection.Input);

                connection.Execute("usp_RegistrarCertificadoEsterilizacion", parameter, commandType: CommandType.StoredProcedure);

                connection.Dispose();
            }
            return si;
        }

        public async Task<int> RegistrarLote(LoteEntity lote)
        {
            int result = 0;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "INSERT INTO TBMCCLote(Descripcion, Lote, Expira, Grupo)" +
                             "VALUES(@Descripcion, @Lote, @Expira, @Identificador)";

                result = await connection.ExecuteAsync(sql, new { lote.Descripcion, lote.Lote, lote.Expira, lote.Identificador });
                connection.Dispose();
            }

            return result;
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

        public async Task<DatosFormatoListarOrdenFabricacionModel> ObtenerInformacionLote(string NumeroLote)
        {

            DatosFormatoListarOrdenFabricacionModel result = new DatosFormatoListarOrdenFabricacionModel();

            string sql = "SELECT FECHAPRODUCCION FechaProduccion, RTRIM(a.ITEM) Item, RTRIM(b.NumeroDeParte) NumeroParte, RTRIM(b.MarcaCodigo) Marca, RTRIM(b.DescripcionLocal) DescripcionLocal, " +
                         "RTRIM(c.NombreCompleto) Cliente, RTRIM(a.NUMEROLOTE) OrdenFabricacion, a.REFERENCIANUMERO Lote, 0 ContraMuestra, RTRIM(SUBSTRING(a.NumeroLotePrincipal,0,CHARINDEX('-',a.NumeroLotePrincipal,0)) )   NumeroCaja , a.AuditableFlag " +
                         "FROM PROD_UNILENE2..EP_PROGRAMACIONLOTE a " +
                         "INNER JOIN PROD_UNILENE2..WH_ItemMast b ON a.ITEM = b.Item " +
                         "INNER JOIN PROD_UNILENE2..PersonaMast c ON a.Cliente = c.Persona " +
                         "WHERE a.ESTADO <> 'AN'  AND a.REFERENCIANUMERO LIKE '"+ NumeroLote +"%' " +
                         "SELECT b.NumeroLote, SUM(b.Cantidad) Calculo , IIF(SUM(b.Cantidad)>0,CAST(1 AS BIT),CAST(0 AS BIT)) Permitir FROM PROD_UNILENE2..EP_PROGRAMACIONLOTE a  " +
                         "RIGHT JOIN TBMKardexInternoCC b ON a.REFERENCIANUMERO = b.NumeroLote " +
                         "WHERE a.ESTADO <> 'AN' AND a.REFERENCIANUMERO LIKE '" + NumeroLote + "%' " +
                         "GROUP BY  b.NumeroLote";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync(sql))
                {
                    result.InformacionLote = result_db.Read<FormatoEstructuraObtenerOrdenFabricacion>().ToList();
                    result.Detalle = result_db.Read<FormatoEstructuraDetalleOrdenFabricacionkardexInterno>().ToList();
                }

                connection.Dispose();
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoListarTransaccion>> ListarTransaccionItem(string NumeroLote, string codAlmacen)
        {
            IEnumerable<DatosFormatoListarTransaccion> result =  new List<DatosFormatoListarTransaccion>();

            string sql ="SELECT  RTRIM (c.Lote) Lote, a.Periodo , RTRIM(CONCAT (a.ReferenciaTipoDocumento,'-',a.ReferenciaNumeroDocumento)) AS Documento, a.Cantidad, RTRIM(a.AlmacenCodigo) AlmacenCodigo, " +
                        "RTRIM(CONCAT(b.ReferenciaTipoDocumento, ' ', b.ReferenciaNumeroDocumento)) AS DocumentoTransaccion " +
                        "FROM PROD_UNILENE2..WH_Kardex  a WITH(NOLOCK) " +
                        "INNER JOIN PROD_UNILENE2..WH_TransaccionHeader b WITH(NOLOCK)ON(a.ReferenciaCompaniaSocio = b.CompaniaSocio AND a.ReferenciaTipoDocumento = b.TipoDocumento AND a.ReferenciaNumeroDocumento = b.NumeroDocumento) " +
                        "INNER JOIN PROD_UNILENE2..WH_TransaccionDetalle c WITH(NOLOCK) ON c.CompaniaSocio = a.ReferenciaCompaniaSocio AND c.TipoDocumento = a.ReferenciaTipoDocumento AND c.NumeroDocumento = a.ReferenciaNumeroDocumento AND c.Secuencia = a.ReferenciaSecuencia " +
                        "INNER JOIN PROD_UNILENE2..WH_ItemMast f WITH(NOLOCK) ON a.Item = f.Item " +
                        "INNER JOIN PROD_UNILENE2..WH_AlmacenMast d WITH(NOLOCK) ON a.AlmacenCodigo = d.AlmacenCodigo " +
                        "LEFT JOIN PROD_UNILENE2..WH_ItemAlmacenLote e WITH(NOLOCK)ON e.Item = a.Item AND e.Condicion = a.Condicion AND e.AlmacenCodigo = a.AlmacenCodigo AND e.Lote = a.Lote WHERE(a.Condicion = '0') " +
                        "AND(d.CompaniaSocio = '01000000')  AND(a.Condicion = '0') AND c.Lote = @NumeroLote AND b.AlmacenCodigo=@codAlmacen " +
                        "ORDER BY a.Periodo ASC, a.Fecha ASC";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {

                result = await context.QueryAsync<DatosFormatoListarTransaccion>(sql, new { NumeroLote, codAlmacen });

            }

            return result;
        }


        public async Task<int> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato, int idUsuario)
        {   
            int result = 1;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {   
                await connection.ExecuteAsync("usp_RegistrarNumeroCaja_ContraMuestra", new { dato.numeroCaja, dato.contraMuestra, dato.fechaProduccion, dato.lote,dato.ordenFabricacion,dato.item, idUsuario }, commandType: CommandType.StoredProcedure);
                
                connection.Dispose();
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoKardexInternoGCM>> ListarKardexInternoNumeroLote(string NumeroLote)
        {
            IEnumerable<DatosFormatoKardexInternoGCM> result = new List<DatosFormatoKardexInternoGCM>();

            string sql = "SELECT Id IdKardex, NumeroLote, OrdenFabricacion, TipoTransaccion, Cantidad, Usuario, Comentarios, Estado  FROM TBMKardexInternoCC WHERE NumeroLote=@NumeroLote";
           

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {

                result = await context.QueryAsync<DatosFormatoKardexInternoGCM>(sql, new { NumeroLote });

            }

            return result;
        }


        public async Task<int> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato, int idUsuario)
        {
            int result = 1;
            string sql2 = "SELECT Usuario FROM  TBMUsuario WHERE CodUsuario = @idUsuario";
            string sql1 = "INSERT INTO TBMKardexInternoCC (NumeroLote,OrdenFabricacion,TipoTransaccion,Cantidad,Usuario,FechaTransaccion,Estado,Comentarios)  " +
                         "VALUES(@Lote, @ordenFabricacion,@Transaccion ,IIF(@Transaccion='NI',(@Cantidad),(@Cantidad*-1)), @Usuario, GETDATE(), 'A', @Comentario); ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string Usuario = await connection.QueryFirstOrDefaultAsync<string>(sql2, new { idUsuario });
                await connection.ExecuteAsync(sql1, new { dato.Lote, dato.OrdenFabricacion,dato.Transaccion,dato.Cantidad, Usuario ,dato.Comentario });

                connection.Dispose();
            }

            return result;
        }


        public async Task<int> ActualizarKardexInternoGCM(int idKardex, string comentarios, int idUsuario)
        {
            int result = 1;

            string sql = "UPDATE TBMKardexInternoCC SET Comentarios=@comentarios ,Usuario=(SELECT Usuario FROM  TBMUsuario WHERE CodUsuario=@idUsuario) WHERE Id=@idKardex ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync(sql, new { idKardex, comentarios, idUsuario });

                connection.Dispose();
            }

            return result;
        }

        public async Task<IEnumerable<FormatoEstructuraObtenerOrdenFabricacion>> ExportarOrdenFabricacionCaja(string anioProduccion)
        {

            IEnumerable<FormatoEstructuraObtenerOrdenFabricacion> result = new List<FormatoEstructuraObtenerOrdenFabricacion>();

            string sql = "SELECT  a.NUMEROLOTE OrdenFabricacion, substring(a.referencianumero,1,8) Lote , FECHAPRODUCCION FechaProduccion ,RTRIM(a.ITEM) Item,RTRIM(b.NumeroDeParte) NumeroParte,RTRIM(b.MarcaCodigo) Marca, RTRIM(b.DescripcionLocal) DescripcionLocal, " +  
                         "RTRIM(c.NombreCompleto) Cliente, cast(a.CANTIDADMUESTRA as DECIMAL(14, 2)) ContraMuestra, RTRIM(a.NumeroLotePrincipal)  NumeroCaja " +
                         "FROM PROD_UNILENE2..EP_PROGRAMACIONLOTE a " +
                         "INNER JOIN PROD_UNILENE2..WH_ItemMast b ON a.ITEM = b.Item " +
                         "INNER JOIN PROD_UNILENE2..PersonaMast c ON a.Cliente = c.Persona " +
                         "WHERE(a.NumeroLotePrincipal IS NOT null  OR a.NumeroLotePrincipal != '') AND a.ESTADO <> 'AN' " +
                         "AND CAST(YEAR(FECHAPRODUCCION) AS varchar) = @anioProduccion " +
                         "ORDER BY a.NumeroLote , a.referencianumero asc";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<FormatoEstructuraObtenerOrdenFabricacion>(sql, new { anioProduccion }) ;
            }

            return result;
        }


        public async Task<(List<DatosFormatosListarControlLotes>, int)> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato)
        {

            (List<DatosFormatosListarControlLotes> DetalleItemControlLotes, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_Listar_informacion_Control_lotes", dato, commandType: CommandType.StoredProcedure))
                {
                    result.DetalleItemControlLotes = result_db.Read<DatosFormatosListarControlLotes>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }
                connection.Dispose();

            }

            return result;
        }

        public async Task<int> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato)
        {
            int result = 1;

            string sql = "  UPDATE PROD_UNILENE2..EP_PROGRAMACIONLOTE SET FechaEntrega =@fechaEntrega WHERE NUMEROLOTE=@ordenFabricacion AND REFERENCIANUMERO=@lote ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync(sql, new { dato.lote, dato.ordenFabricacion, dato.fechaEntrega });

                connection.Dispose();
            }

            return result;
        }


    }
}
