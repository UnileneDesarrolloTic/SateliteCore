using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ProduccionRepository : IProduccionRepository
    {
        private readonly IAppConfig _appConfig;

        public ProduccionRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<SeguimientoProductoArimaModel> SeguimientoProductosArima(string periodo)
        {
            SeguimientoProductoArimaModel result = new SeguimientoProductoArimaModel();

            using (SqlConnection springContext = new SqlConnection(_appConfig.contextSpring))
            {
                using SqlMapper.GridReader multi = await springContext.QueryMultipleAsync("usp_Satelite_ProductoTerminadoArima", new { periodo }, commandType: CommandType.StoredProcedure);
                result.Productos = multi.Read<ProductoArimaModel>().ToList();
                result.DetalleTransito = multi.Read<TransitoProductoArimaModel>().ToList();
            }

            return result;
        }

        public async Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla)
        {
            SeguimientoCandMPAGenericModel result = new SeguimientoCandMPAGenericModel();

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                result.SeguimientoCandidatosMPA = await satelliteContext.QueryAsync<SeguimientoCandMPAModel>("usp_pro_SeguimientoCandidatoMPA", new { regla }, commandType: CommandType.StoredProcedure);
                result.OrdenComprasPendientes = await satelliteContext.QueryAsync<DetalleSeguimientoCandMPAModel>("usp_pro_SeguimientoDetalleCandidatoMPA", new { regla }, commandType: CommandType.StoredProcedure);
                result.DetalleTotalesProducto = await satelliteContext.QueryAsync<TotalesProductoMPArimaModel>("usp_pro_TotalesCanditosMPA", commandType: CommandType.StoredProcedure);
                satelliteContext.Dispose();
            }

            return result;
        }

        public async Task<IEnumerable<SeguimientoCandMPAModel>> ExportarAgujasMateriaPrima(string regla)
        {
            IEnumerable<SeguimientoCandMPAModel> result = new List<SeguimientoCandMPAModel>();

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                result = await satelliteContext.QueryAsync<SeguimientoCandMPAModel>("usp_pro_SeguimientoCandidatoMPA", new { regla }, commandType: CommandType.StoredProcedure);
              
            }

            return result;
        }


        public async Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item)
        {
            List<DetalleControlCalidadItemMP> result;

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                using (var result_db = await satelliteContext.QueryMultipleAsync("usp_pro_DetalleControlCalidadMP", new { Item }, commandType: CommandType.StoredProcedure))
                {
                    result = result_db.Read<DetalleControlCalidadItemMP>().ToList();
                } 
                satelliteContext.Dispose();
            }

            return result;
        }

        public async Task<bool> MostrarColumnaMP(int Usuario)
        {
            bool result;

            using (var satelliteContext = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                 using (var result_db = await satelliteContext.QueryMultipleAsync("usp_pro_MostrarColumnaPorUsuarioMP", new { Usuario }, commandType: CommandType.StoredProcedure))
                 {
                     result = result_db.Read<bool>().First();
                 }
                 satelliteContext.Dispose();
                
            }

            return result;
        }


        public async Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro)
        {

            (IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros) result;
            DynamicParameters parametros = new DynamicParameters();

            parametros.Add("@FechaInicio", filtro.FechaInicio);
            parametros.Add("@FechaFin", filtro.FechaFin);
            parametros.Add("@Item", filtro.Item);
            parametros.Add("@Pagina", filtro.Pagina);
            parametros.Add("@RegistrosPorPagina", filtro.RegistrosPorPagina);
            parametros.Add("@TotalRegistos", dbType: DbType.Int32, direction: ParameterDirection.Output);

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                result.ListaPedidos = await satelliteContext.QueryAsync<PedidosCreadosAutoLogModel>("usp_Pro_ListarPedidosCreadosAuto", parametros, commandType: CommandType.StoredProcedure);
                result.TotalRegistros = parametros.Get<int>("@TotalRegistos");
                satelliteContext.Dispose();
            }

            return result;
        }

        public async Task<SeguimientoComprasMPArima> SeguimientoCompraMPArima(PronosticoCompraMP dato)
        {
            SeguimientoComprasMPArima result = new SeguimientoComprasMPArima();

            using (SqlConnection DMVentasContext = new SqlConnection(_appConfig.contextSpring))
            {
                using SqlMapper.GridReader multi = await DMVentasContext.QueryMultipleAsync("usp_Satelite_CompraMateriaPrimaArima", dato, commandType: CommandType.StoredProcedure);
                result.Productos = multi.Read<CompraMPArimaModel>().ToList();
                result.DetalleTransito = multi.Read<DCompraMPArimaModel>().ToList();
                result.DetalleCalidad = multi.Read<CompraMPArimaDetalleControlCalidad>().ToList();

            }

            return result;
        }


        public async Task<FormatoEstructuraLoteEtiquetas> LoteFabricacionEtiquetas(string NumeroLote)
        {
            FormatoEstructuraLoteEtiquetas result = new FormatoEstructuraLoteEtiquetas();

            string sql = "SELECT FECHAPRODUCCION FechaProduccion ,RTRIM(a.ITEM) Item,RTRIM(b.NumeroDeParte) NumeroParte,RTRIM(b.MarcaCodigo) Marca, RTRIM(b.DescripcionLocal) DescripcionLocal,  " +
                         "RTRIM(c.NombreCompleto) Cliente,RTRIM(a.NUMEROLOTE) OrdenFabricacion, SUBSTRING(RTRIM(a.REFERENCIANUMERO), 1, 8) Lote  , a.transferidoflag " +
                         "FROM PROD_UNILENE2..EP_PROGRAMACIONLOTE a " +
                         "INNER JOIN PROD_UNILENE2..WH_ItemMast b ON a.ITEM = b.Item " +
                         "INNER JOIN PROD_UNILENE2..PersonaMast c ON a.Cliente = c.Persona " +
                         "WHERE a.REFERENCIANUMERO = @NumeroLote AND a.ESTADO <> 'AN' ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {

                result = await context.QueryFirstOrDefaultAsync<FormatoEstructuraLoteEtiquetas>(sql, new { NumeroLote });

            }

            return result;
        }

        public async Task<int> RegistrarLoteFabricacionEtiquetas(List<DatosEstructuraLoteEtiquetasModel> dato, int idUsuario)
        {
            int result = 1;
            string sql = "INSERT INTO TBMPRLoteEstado (Lote,OrdenFabricacion,FechaRegistro,Estado,Usuario) VALUES (@lote,@ordenFabricacion,GETDATE(),'A',@idUsuario) " +
                          "UPDATE PROD_UNILENE2..EP_PROGRAMACIONLOTE SET transferidoflag='N' WHERE REFERENCIANUMERO = @lote";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                foreach (DatosEstructuraLoteEtiquetasModel item in dato)
                {
                    await connection.ExecuteAsync(sql, new { item.lote, item.ordenFabricacion, idUsuario });
                }
                connection.Dispose();
            }

            return result;
        }


        public async Task<IEnumerable<DatoFormatoLoteEstado>> ListarLoteEstado()
        {
            IEnumerable<DatoFormatoLoteEstado> result = new List<DatoFormatoLoteEstado>();

            string sql = "select a.Id, a.Lote,a.OrdenFabricacion,FechaRegistro ,a.Estado, RTRIM(b.CodigoUsuario) Usuario FROM TBMPRLoteEstado a " +
                         "LEFT JOIN PROD_UNILENE2..Empleadomast b ON a.Usuario = b.Empleado " +
                         "WHERE a.Estado <> 'I'";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {

                result = await context.QueryAsync<DatoFormatoLoteEstado>(sql);

            }
            return result;
        }


        public async Task<int> ModificarLoteEstado(DatosFormatoRequestLoteEstado dato)
        {
            int result = 1;
            string sql = "UPDATE TBMPRLoteEstado set Estado='I' where id=@id " +
                         "UPDATE PROD_UNILENE2..EP_PROGRAMACIONLOTE SET TRANSFERIDOFLAG = 'S' where REFERENCIANUMERO = @lote";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                
                    await connection.ExecuteAsync(sql, new { dato.id, dato.lote });
              
                connection.Dispose();
            }

            return result;
        }


        public async Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Origen, string Anio, string Regla)
        {
            DatosFormatoInformacionCalendarioSeguimientoOC result = new DatosFormatoInformacionCalendarioSeguimientoOC();

            using (SqlConnection connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await connection.QueryMultipleAsync("usp_Listar_item_Seguimiento_Compra", new { Origen, Anio, Regla }, commandType: CommandType.StoredProcedure);
                result.Calendario = multi.Read<DatosFormatoCalentarioSeguimientoOC>().ToList();
                result.DetalleCalendario = multi.Read<DatosFormatoDetalleCalendarioSeguimientoOC>().ToList();

            }
            return result;
        }

        public async Task<DatosFormatoInformacionItemOrdenCompra> BuscarItemOrdenCompra(string Item,string Anio)
        {
            DatosFormatoInformacionItemOrdenCompra result = new DatosFormatoInformacionItemOrdenCompra();

            using (SqlConnection connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await connection.QueryMultipleAsync("usp_Buscar_Item_Orden_Compra", new { Item, Anio }, commandType: CommandType.StoredProcedure);
                result.informacionItem = multi.Read<FormatoDatoInformacionItem>().ToList();
                result.ListaOrdenCompra = multi.Read<FormatoDatosOrdenCompraItem>().ToList();
                result.Detalle = multi.Read<object>().ToList();

            }

            return result;
        }

        public async Task<int> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra dato)
        {
            int result=0;
            string script = "UPDATE PROD_UNILENE2..WH_OrdenCompradetalle SET FechaPrometida=@fechaPrometida WHERE NumeroOrden=@numeroOrden  AND Item=@item ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(script, new { dato.fechaPrometida, dato.item , dato.numeroOrden});
            }

            return result;
        }

        public async Task<(object cabecera, object detalle)> VisualizarOrdenCompra(string OrdenCompra)
        {
            (object cabecera, object detalle) datosOrdenCompra;

            string script = "SELECT RTRIM(a.NumeroOrden) NumeroOrden,RTRIM(b.Busqueda) Proveedor,FechaPreparacion, FechaPrometida, FechaEnvioProveedor ,a.Estado " +
                            "FROM PROD_UNILENE2..WH_OrdenCompra a INNER JOIN PROD_UNILENE2..PersonaMast b ON a.Proveedor = b.Persona " +
                            "WHERE NumeroOrden = @OrdenCompra  " +
                            "SELECT RTRIM(NumeroOrden) NumeroOrden,RTRIM(Item) Item, RTRIM(Descripcion) Descripcion, RTRIM(UnidadCodigo) UnidadCodigo, CantidadPedida , CantidadRecibida , Estado , FechaPrometida , CAST(0 AS BIT) isSelected   " +
                            "FROM PROD_UNILENE2..WH_OrdenCompradetalle WHERE NumeroOrden = @OrdenCompra ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using var result =  await context.QueryMultipleAsync(script, new { OrdenCompra });
                datosOrdenCompra.cabecera = result.Read().FirstOrDefault();
                datosOrdenCompra.detalle = result.Read().ToList();
            }

            return datosOrdenCompra;
        }

        public async Task<int> ActualizarFechaComprometidaMasiva(DatosFormatoCabeceraOrdenCompraModel dato)
        {
            int result = 0;
            
            string script = "UPDATE PROD_UNILENE2..WH_OrdenCompradetalle SET FechaPrometida=@FechaLlegada WHERE NumeroOrden=@NumeroOrden  AND Item=@Item AND estado<>'CO' ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
               
               foreach(DatosFormatoDetalleOrdenCompraMasivo item in dato.Detalle)
                {
                    await context.ExecuteAsync(script, new { item.NumeroOrden, item.Item, dato.FechaLlegada });
                }
               
            }

            return result;

        }

        public async Task<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>> SeguimientoOCDrogueria(int idproveedor)
        {
            IEnumerable<DatosFormatoReporteSeguimientoDrogueria> result = new List<DatosFormatoReporteSeguimientoDrogueria>();
            
            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
               result = await context.QueryAsync<DatosFormatoReporteSeguimientoDrogueria>("usp_Seguimiento_OCdrogueria",new { idproveedor }, commandType: CommandType.StoredProcedure);
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>> MostrarOrdenCompraDrogueria(string Item)
        {
            IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria> result = new List<DatosFormatoMostrarOrdenCompraDrogueria>();

          
            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatoMostrarOrdenCompraDrogueria>("usp_MostrarOdenCompraDrogueria", new {Item } , commandType: CommandType.StoredProcedure);
            }
            return result;
        }


        public async Task<IEnumerable<DatosFormatoMostrarProveedorDrogueria>> MostrarProveedorDrogueria()
        {
            IEnumerable<DatosFormatoMostrarProveedorDrogueria> result = new List<DatosFormatoMostrarProveedorDrogueria>();

            string sql = "SELECT Id IdProveedor, Descripcion NombreProveedor FROM TBMConfiguracion " +
                         "WHERE Observaciones='PROVEEDORDROGUERIA' AND Estado = 'A' ;";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoMostrarProveedorDrogueria>(sql);
            }
            return result;
        }

    }
}
