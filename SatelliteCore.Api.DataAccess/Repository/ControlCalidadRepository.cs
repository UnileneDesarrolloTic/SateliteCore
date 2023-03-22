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
                        "AND(d.CompaniaSocio = '01000000')  AND(a.Condicion = '0') AND(a.AlmacenCodigo = @codAlmacen) AND((CASE WHEN f.ItemTipo = 'PT' THEN e.LoteFabricacion ELSE e.Lote END) = @NumeroLote) AND(a.Condicion = '0') " +
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
                         "RTRIM(c.NombreCompleto) Cliente, (SELECT SUM(CANTIDAD) FROM TBMKardexInternoCC WHERE NUMEROLOTE=a.REFERENCIANUMERO  AND ORDENFABRICACION=a.NUMEROLOTE  AND ESTADO='A') ContraMuestra," +
                         "(SELECT TOP 1 FechaTransaccion FROM TBMKardexInternoCC WHERE NUMEROLOTE=a.REFERENCIANUMERO  AND ORDENFABRICACION=a.NUMEROLOTE  AND ESTADO='A' AND TipoTransaccion='NI')  FechaIngreso , RTRIM(a.NumeroLotePrincipal)  NumeroCaja " +
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


        public async Task<IEnumerable<DatosFormatosListarControlLotes>> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato)
        {
            IEnumerable<DatosFormatosListarControlLotes> result = new List<DatosFormatosListarControlLotes>();


            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatosListarControlLotes>("usp_Listar_informacion_Control_lotes", dato, commandType: CommandType.StoredProcedure);
               
                connection.Dispose();

            }

            return result;
        }

        public async Task<int> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato)
        {
            int result = 1;

            string sql = "  UPDATE PROD_UNILENE2..EP_PROGRAMACIONLOTE SET FechaEntrega =@fechaEntrega , Proyecto =@destruible , COMENTARIOS=@comentarios   WHERE NUMEROLOTE=@ordenFabricacion AND REFERENCIANUMERO=@lote ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync(sql, new { dato.lote, dato.ordenFabricacion, dato.fechaEntrega, dato.destruible , dato.comentarios });

                connection.Dispose();
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoTablaNumerodeParte>> ListarMaestroNumeroParte(string Grupo,string Tabla)
        {
            IEnumerable<DatosFormatoTablaNumerodeParte> result = new List<DatosFormatoTablaNumerodeParte>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoTablaNumerodeParte>("SP_UNILENE_LO_MAESTRO_NUMERO_PARTE", new { GRUPO= Grupo, TABLA=Tabla }, commandType: CommandType.StoredProcedure);
                
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoTablaAbributoModel>> ListarAtributos()
        {
           
            IEnumerable<DatosFormatoTablaAbributoModel> result = new List<DatosFormatoTablaAbributoModel>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoTablaAbributoModel>("SP_LISTAR_TPRO_CANTIDAD_AGUJAS", commandType: CommandType.StoredProcedure);

            }
            return result ;
        }

        public async Task<IEnumerable<DatosFormatoTablaDescripcionModel>> ListarDescripcion(string Marca,string Hebra)
        {

            IEnumerable<DatosFormatoTablaDescripcionModel> result = new List<DatosFormatoTablaDescripcionModel>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoTablaDescripcionModel>("SP_LISTAR_TPRO_DESCRIPCION", new { ID_MARCA = Marca , ID_HEBRA=Hebra } , commandType: CommandType.StoredProcedure);

            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoTablaLeyendaModel>> ListarLeyenda(string Marca, string Hebra)
        {

            IEnumerable<DatosFormatoTablaLeyendaModel> result = new List<DatosFormatoTablaLeyendaModel>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoTablaLeyendaModel>("SP_LISTAR_TPRO_LEYENDA", new { ID_MARCA = Marca, ID_HEBRA = Hebra }, commandType: CommandType.StoredProcedure);

            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoTablaPruebasModel>> ListarTablaPrueba(string Metodologia)
        {

            IEnumerable<DatosFormatoTablaPruebasModel> result = new List<DatosFormatoTablaPruebasModel>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoTablaPruebasModel>("SP_LISTAR_TPRO_PRUEBA", new { Metodologia }, commandType: CommandType.StoredProcedure);

            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel>> ListarObtenerAgujasDescripcionNuevo()
        {

            IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel> result = new List<DatosFormatoObtenerTablaAgujasNuevoModel>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoObtenerTablaAgujasNuevoModel>("SP_LISTAR_TPRO_CANTIDAD_AGUJAS", commandType: CommandType.StoredProcedure);

            }
            return result;
        }

      

        public async Task<int> NuevoDescripcionDT(DatosFormatoActualizacionDescripcionModel dato,string idUsuario)
        {

            int IdDescripcion = 0;

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                IdDescripcion = await connection.QuerySingleAsync<int>("SP_INSERTAR_TPRO_DESCRIPCION", new { ID_MARCA = dato.Marca, ID_HEBRA = dato.Hebra, DESCRIPCIONLOCAL = dato.DescripcionLocal, DESCRIPCIONINGLES = dato.DescripcionIngles, USUARIO = idUsuario }, commandType: CommandType.StoredProcedure);
                foreach (DatosFormatoDetalleAgujaDescripcion valor in dato.DetalleAgujas)
                {   
                    if(valor.descripcionlocal.Trim()!="" || valor.descripcionlocal.Trim() != "")
                       await connection.ExecuteAsync("SP_INSERTAR_TPRO_CARACTERISTICA_DESCRIPCION", new { ID_DESCRIPCION = IdDescripcion, ID_CANTIDAD_AGUJA=valor.iD_AGUJA, DESCRIPCIONLOCAL=valor.descripcionlocal, DESCRIPCIONINGLES=valor.descripcioningles }, commandType: CommandType.StoredProcedure);
                }
            }
            return IdDescripcion;
        }

        public async Task<int> EliminarDescripcionDT(string IdDescripcion)
        {

            int result = 0;
            string sql = "UPDATE TPRO_DESCRIPCION SET Estado='I' WHERE ID_DESCRIPCION=@IdDescripcion ";

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
               await connection.ExecuteAsync(sql,new {IdDescripcion});

            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoObtenerAgujasDescripcionModel>> ListarObtenerAgujasDescripcionActualizar(string IdDescripcion)
        {

            IEnumerable<DatosFormatoObtenerAgujasDescripcionModel> result = new List<DatosFormatoObtenerAgujasDescripcionModel>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoObtenerAgujasDescripcionModel>("SP_LISTAR_TPRO_CARACTERISTICAS_DESCRIPCION", new { ID_DESCRIPCION = IdDescripcion }, commandType: CommandType.StoredProcedure);

            }
            return result;
        }

        public async Task<int> ActualizarDescripcionDT(DatosFormatoActualizacionDescripcionModel dato, string Usuario)
        {

            int result = 0;

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await connection.ExecuteAsync("SP_ACTUALIZAR_TPRO_DESCRIPCION", new { ID_DESCRIPCION = dato.IdDescripcion, ID_MARCA = dato.Marca, ID_HEBRA = dato.Hebra, DESCRIPCIONLOCAL = dato.DescripcionLocal, DESCRIPCIONINGLES = dato.DescripcionIngles, USUARIO = Usuario }, commandType: CommandType.StoredProcedure);
                foreach (DatosFormatoDetalleAgujaDescripcion valor in dato.DetalleAgujas)
                {
                    if (valor.descripcionlocal.Trim() != "" || valor.descripcionlocal.Trim() != "")
                        await connection.ExecuteAsync("SP_INSERTAR_TPRO_CARACTERISTICA_DESCRIPCION", new { ID_DESCRIPCION = dato.IdDescripcion, ID_CANTIDAD_AGUJA = valor.iD_AGUJA, DESCRIPCIONLOCAL = valor.descripcionlocal, DESCRIPCIONINGLES = valor.descripcioningles }, commandType: CommandType.StoredProcedure);
                }

            }
            return result;
        }

        public async Task<int> RegistrarActualizarLeyendaDT(DatosFormatoLeyendaDTModel dato, string Usuario)
        {

            int result = 0;

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                if (dato.IdLeyenda == 0)
                {
                    await connection.ExecuteAsync("SP_INSERTAR_TPRO_LEYENDA", new { NUM_REGISTRO = dato.RegistroSanitario, ID_MARCA = dato.Marca, ID_HEBRA = dato.Hebra, TECNICA = dato.TecnicaEspaniol, METODO = dato.MetodoEspaniol, DETALLE = dato.DetalleEspaniol, TECNICA_INGLES = dato.TecnicaIngles, METODO_INGLES = dato.MetodoIngles, DETALLE_INGLES = dato.DetalleIngles, USUARIO = Usuario }, commandType: CommandType.StoredProcedure);
                }
                else
                {
                    await connection.ExecuteAsync("SP_ACTUALIZAR_TPRO_LEYENDA", new { ID_LEYENDA=dato.IdLeyenda, NUM_REGISTRO = dato.RegistroSanitario, ID_MARCA = dato.Marca, ID_HEBRA = dato.Hebra, TECNICA = dato.TecnicaEspaniol, METODO = dato.MetodoEspaniol, DETALLE = dato.DetalleEspaniol, TECNICA_INGLES = dato.TecnicaIngles, METODO_INGLES = dato.MetodoIngles, DETALLE_INGLES = dato.DetalleIngles, USUARIO = Usuario }, commandType: CommandType.StoredProcedure);
                }

            }
            return result;
        }


        public async Task<int> EliminarLeyendaDT(string IdLeyenda)
        {

            int result = 0;
            string sql = "UPDATE TPRO_LEYENDA SET Estado='I' WHERE ID_LEYENDA=@IdLeyenda ";

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await connection.ExecuteAsync(sql, new { IdLeyenda });
            }

            return result;
        }


        public async Task<int> RegistrarActualizarPruebaDT(DatosFormatoNuevoPruebaModel dato, string idUsuario)
        {

            int result = 0;


            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
               await connection.ExecuteAsync("SP_INSERTAR_TPRO_PRUEBA", new { ID_AGRUPADOR_HEBRA = dato.IdAgrupadoHebra,
                                                                              ID_METODOLOGIA = dato.IdMedologia, 
                                                                              CALIBRE_USP = dato.IdCalibre, 
                                                                              DESCRIPCIONLOCAL = dato.DescripcionLocal, 
                                                                              DESCRIPCIONINGLES = dato.DescripcionIngle,
                                                                              UNIDAD_MEDIDA = dato.IdUnidadMedida, 
                                                                              ESPECIFFICACIONLOCAL = dato.EspecificacionLocal, 
                                                                              ESPECIFFICACIONINGLES = dato.EspecificacionIngles, 
                                                                              VALOR = dato.valor, 
                                                                              USUARIO = idUsuario }, commandType: CommandType.StoredProcedure);

            }
            return result;
        }

        public async Task<int> EliminarPruebaDT(string IdPrueba)
        {

            int result = 0;
            string sql = "UPDATE TPRO_PRUEBA_DETALLE SET Estado='I' WHERE ID_PRUEBA=@IdPrueba " +
                         "UPDATE TPRO_PRUEBA SET Estado='I' WHERE ID_PRUEBA=@IdPrueba";

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await connection.ExecuteAsync(sql, new { IdPrueba });
            }

            return result;
        }


        public async Task<DatosFormatoNumeroLoteProtocoloModel> BuscarNumeroLoteProtocolo(string NumeroLote,string Idioma)
        {
            DatosFormatoNumeroLoteProtocoloModel result = new DatosFormatoNumeroLoteProtocoloModel();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryFirstOrDefaultAsync<DatosFormatoNumeroLoteProtocoloModel>("usp_listar_tpro_protocolo_cabecera", new { NumeroLote , Idioma }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatosDatoListarPruebaProtocolo>> BuscarPruebaFormatoProtocolo(string NumeroLote, string NumeroParte, string Idioma )
        {
            IEnumerable<DatosFormatosDatoListarPruebaProtocolo> result = new List<DatosFormatosDatoListarPruebaProtocolo>();
             

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatosDatoListarPruebaProtocolo>("usp_lista_formato_protocolo", new {  NumeroParte, NumeroLote , Idioma }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<int> RegistrarControlProcesoProtocolo(DatosFormatoControlProcesosProtocoloModel dato, string Usuario)
        {
            int ID = 0;
            int contarA = 1;
            int contarB = 1;
            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                ID = await connection.QueryFirstOrDefaultAsync<int>("SP_INSERTAR_TPRO_RESULTADO_CABECERA", new { LOTE = dato.Numerolote, USUARIO = Usuario, FECHA_ANALISIS = dato.fechaanalisis }, commandType: CommandType.StoredProcedure);
                
                foreach (DatosFormatosTablaAControlProcesos item in dato.TablaLongitud)
                {
                    await connection.ExecuteAsync("SP_INSERTAR_TPRO_RESULTADO_DETALLE", new { ID_CABECERA = ID, TABLA = 'A', SECUENCIA = contarA , COL_1=item.LongitudD  , COL_2=item.DiametroD }, commandType: CommandType.StoredProcedure);
                    contarA++;
                }
        
                foreach (DatosFormatosTablaBControlProcesos item in dato.TablaResistencia)
                {
                    await connection.ExecuteAsync("SP_INSERTAR_TPRO_RESULTADO_DETALLE", new { ID_CABECERA = ID, TABLA = 'B', SECUENCIA = contarB, COL_1 = item.TensionNewtons, COL_2 = item.AgujasNewtons }, commandType: CommandType.StoredProcedure);
                    contarB++;
                }
            }

            return ID;
        }


        public async Task<int> RegistrarControlPTProtocolo(DatosFormatoControlProductoTermino dato, string Usuario)
        {
            int ID = 0;
            int contarA = 1;
            int contarB = 1;
            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                ID = await connection.QueryFirstOrDefaultAsync<int>("SP_INSERTAR_TPRO_RESULTADO_CABECERA", new { LOTE = dato.Numerolote, USUARIO = Usuario, FECHA_ANALISIS = dato.fechaanalisis }, commandType: CommandType.StoredProcedure);

                foreach (DatosFormatosTablaAControlProcesos item in dato.TablaLongitud)
                {
                    await connection.ExecuteAsync("SP_INSERTAR_TPRO_RESULTADO_DETALLE", new { ID_CABECERA = ID, TABLA = 'C', SECUENCIA = contarA, COL_1 = item.LongitudD, COL_2 = item.DiametroD }, commandType: CommandType.StoredProcedure);
                    contarA++;
                }

                foreach (DatosFormatosTablaBControlProcesos item in dato.TablaResistencia)
                {
                    await connection.ExecuteAsync("SP_INSERTAR_TPRO_RESULTADO_DETALLE", new { ID_CABECERA = ID, TABLA = 'D', SECUENCIA = contarB, COL_1 = item.TensionNewtons, COL_2 = item.AgujasNewtons }, commandType: CommandType.StoredProcedure);
                    contarB++;
                }
            }

            return ID;
        }

        public async Task<int> RegistrarPruebasEfectuadasProtocolo(DatosFormatoPruebasEfectuasProtocolos dato, string idUsuario)
        {
            int ID = 0;
            int contarB = 1;
            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                foreach (DatosFormatoDetallePruebasProtocolos item in dato.TablaPrueba)
                {
                    await connection.ExecuteAsync("SP_INSERTAR_TPRO_RESULTADO_DETALLE", commandType: CommandType.StoredProcedure);
                    contarB++;
                }
            }

            return ID;
        }

        public async Task<IEnumerable<DatosFormatoInformacionResultadoProtocolo>> BuscarInformacionResultadoProtocolo(string NumeroLote)
        {
            IEnumerable<DatosFormatoInformacionResultadoProtocolo> result = new List<DatosFormatoInformacionResultadoProtocolo>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoInformacionResultadoProtocolo>("SP_LISTAR_PRUEBA_RESULTADO", new { LOTE = NumeroLote }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<int> InsertarCabeceraFormatoProtocolo(DatosFormatoCabeceraFormatoProtocolo dato, string UsuarioSesion)
        {
            int result = 1;
            int ID = 0;
            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                ID = await connection.QueryFirstOrDefaultAsync<int>("SP_INSERTAR_FORMATO_PROTOCOLO_CABECERA", 
                        new { ID_IDIOMA = dato.Idioma, FECHA_ANALISIS=dato.fechaanalisis, LOTE=dato.NumeroLote, NUMERODEPARTE=dato.NumeroParte, TECNICA= dato.Tecnica, METODO=dato.Metodo, DETALLE=dato.Detalle, USUARIO= UsuarioSesion }, commandType: CommandType.StoredProcedure);
                
                foreach(DatosFormatoDetalleFormatoProtocolo valor in dato.TablaPrueba)
                {
                    await connection.QueryAsync("SP_INSERTAR_FORMATO_PROTOCOLO_DETALLE",
                        new { ID_FORMATO = ID, DESCRIPCION_PRUEBA = valor.descripcionLocal, UNIDAD_MEDIDA = valor.unidadMedida, ESPECIFICACION = valor.especificacion, VALOR = valor.valor , RESULTADO =valor.resultado, DESCRIPCION_METODOLOGIA = valor.metodologia, ORDEN =  valor.orden}, commandType: CommandType.StoredProcedure);

                }
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoInformacionResultadoProtocolo>> ImprimirControlProceso(string NumeroLote)
        {
            IEnumerable<DatosFormatoInformacionResultadoProtocolo> result = new List<DatosFormatoInformacionResultadoProtocolo>();

            using (var connection = new SqlConnection(_appConfig.ContextUReporteador))
            {
                result = await connection.QueryAsync<DatosFormatoInformacionResultadoProtocolo>("SP_LISTAR_PRUEBA_RESULTADO", new { LOTE = NumeroLote }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoProtocoloPruebaModel>> ImprimirDocumentoProtocolo(string NumeroLote, string Idioma)
        {
            IEnumerable<DatosFormatoProtocoloPruebaModel> result = new List<DatosFormatoProtocoloPruebaModel>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatoProtocoloPruebaModel>("usp_reporte_formato_protocolo", new { NumeroLote  , Idioma }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<ParametroMastEntity> ProtocoloRevisionTerminado()
        {
            ParametroMastEntity result = new ParametroMastEntity();

            string sql = "SELECT CompaniaCodigo, AplicacionCodigo, ParametroClave, DescripcionParametro, Explicacion, TipodeDatoFlag, RTRIM(Texto) Texto, Numero, Fecha, FinanceComunFlag,  Estado,  UltimoUsuario, UltimaFechaModif, ExplicacionAdicional,  Texto1  FROM ParametrosMast WHERE parametroclave='FORMISOPT'"; 

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                result = await connection.QueryFirstOrDefaultAsync<ParametroMastEntity>(sql);
            }

            return result;
        }

        public async Task<string> VersionProtocolo()
        {
            string versionprotocolo = "";
            string sql = "SELECT Explicacion FROM ParametrosMast WHERE ParametroClave = 'FORMISOPT' AND AplicacionCodigo = 'WH' AND CompaniaCodigo = '010000'";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                versionprotocolo = await connection.QueryFirstOrDefaultAsync<string>(sql);
            }

            return versionprotocolo;
        }






    }
}
