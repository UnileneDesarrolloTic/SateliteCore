using Dapper;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class CommonRepository : ICommonRepository
    {
        private readonly IAppConfig _appConfig;

        public CommonRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<TipoDocumentoIdentidadEntity>> ListarTipoDocumentoIndentidad()
        {
            IEnumerable<TipoDocumentoIdentidadEntity> lista = new List<TipoDocumentoIdentidadEntity>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string script = "SELECT Codigo, Descripcion, Abreviatura, Longitud, FlagLongExacta FROM TBMTipoDocumentoIdentidad";
                lista = await connection.QueryAsync<TipoDocumentoIdentidadEntity>(script);

                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<PaisEntity>> ListarPaises()
        {
            IEnumerable<PaisEntity> lista = new List<PaisEntity>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string script = "SELECT Codigo, Nombre, Moneda, Imagen, GentilicioMasculino, GentilicioFemenino FROM TBMPais";
                lista = await connection.QueryAsync<PaisEntity>(script);

                connection.Dispose();
            }

            return lista;
        }

        public async Task<List<MenuEntity>> ListarMenuxUsuario(int usuario)
        {
            List<MenuEntity> lista = new List<MenuEntity>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = (List<MenuEntity>)await connection.QueryAsync<MenuEntity>("usp_ObtenerMenuSatelitexUsuario",
                            new { usuario }, commandType: CommandType.StoredProcedure);

                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<RolEntity>> ListarRoles(string estado)
        {
            IEnumerable<RolEntity> listaRol = new List<RolEntity>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                listaRol = await connection.QueryAsync<RolEntity>(
                        "SELECT CodRol, Titulo, Descripcion, Estado FROM TBMRol WHERE Estado = IIF(@estado = 'T', Estado, @estado)",
                        new { estado });
                connection.Close();
            }

            return listaRol;
        }

        public async Task<List<FamiliaMP>> ListarFamiliaMP(string tipo)
        {
            List<FamiliaMP> lista = new List<FamiliaMP>();
            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = (List<FamiliaMP>)await connection.QueryAsync<FamiliaMP>("SELECT Codigo, Valor1 FROM TBDCatalogo WHERE Estado='A' AND CatalogoID=6  AND RTRIM(PadreCodigo)=@tipo ", new { tipo });
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<ConfiguracionEntity>> ObtenerConfiguracionesSistema(int idConfiguracion, string grupo)
        {
            IEnumerable<ConfiguracionEntity> lista = new List<ConfiguracionEntity>();

            string query = "SELECT IdConfiguracion, Id, Grupo, ValorTexto1, ValorTexto2, ValorEntero1, ValorEntero2, ValorEntero3, ValorDecimal1, ValorDecimal2, " +
                "ValorDecimal3, ValorFecha1, ValorFecha2, ValorFecha3, ValorBit, Estado FROM TBDConfiguracion " +
                "WHERE IdConfiguracion = @idConfiguracion AND Grupo = @grupo AND Estado = 'A' ORDER BY Id ASC";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<ConfiguracionEntity>(query, new { idConfiguracion, grupo });
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<AgrupadorEntity>> ListarAgrupador()
        {
            IEnumerable<AgrupadorEntity> lista = new List<AgrupadorEntity>();

            string query = "SELECT CodigoAgrupador , DescripcionAgrupador , DescripcionCompleta  FROM T_AGRUPADOR WHERE FECHAFINVIG='2100-01-01'";

            using (var connection = new SqlConnection(_appConfig.ContextDMVentas))
            {
                lista = await connection.QueryAsync<AgrupadorEntity>(query);
                connection.Dispose();
            }

            return lista;
        }
        public async Task<IEnumerable<SubAgrupadorEntity>> ListarSubAgrupador(string idAgrupador)
        {
            IEnumerable<SubAgrupadorEntity> lista = new List<SubAgrupadorEntity>();

            string query = "SELECT RTRIM(SUBAGRUPADOR_CODIGO) codSubAgrupador , RTRIM(SUBAGRUPADOR_NOMBRE) NombreSubAgrupador , RTRIM(AGRUPADOR_CODIGO) codAgrupador  FROM T_SUBAGRUPADOR WHERE AGRUPADOR_CODIGO=@idAgrupador";

            using (var connection = new SqlConnection(_appConfig.ContextDMVentas))
            {
                lista = await connection.QueryAsync<SubAgrupadorEntity>(query,new { idAgrupador });
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<LineaEntity>> ListarLinea()
        {
            IEnumerable<LineaEntity> lista = new List<LineaEntity>();

            string query = "SELECT RTRIM(Linea) Linea, RTRIM(DescripcionLocal) DescripcionLocal, RTRIM(DescripcionIngles)  DescripcionIngles FROM WH_ClaseLinea";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await connection.QueryAsync<LineaEntity>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<FamiliaMaestroItemsModel>> ListarFamilia(string idlinea)
        {
            IEnumerable<FamiliaMaestroItemsModel> lista = new List<FamiliaMaestroItemsModel>();

            string query = "SELECT RTRIM(Familia) Familia, RTRIM(DescripcionLocal) DescripcionLocal, RTRIM(DescripcionCompleta) DescripcionCompleta FROM  WH_CLASEFAMILIA WHERE DescripcionLocal like '%sutura%' and Linea = @idlinea AND Estado = 'A' ";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await connection.QueryAsync<FamiliaMaestroItemsModel>(query, new { idlinea });
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<FamiliaMaestroItemsModel>> ListarFamiliaGeneral(string idlinea)
        {
            IEnumerable<FamiliaMaestroItemsModel> lista = new List<FamiliaMaestroItemsModel>();

            string query = "SELECT RTRIM(Familia) Familia, RTRIM(DescripcionLocal) DescripcionLocal, RTRIM(DescripcionCompleta) DescripcionCompleta FROM  WH_CLASEFAMILIA WHERE Linea = @idlinea AND Estado = 'A' ";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await connection.QueryAsync<FamiliaMaestroItemsModel>(query, new { idlinea });
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<SubFamiliaEntity>> ListarSubFamilia(string idlinea, string idfamilia)
        {
            IEnumerable<SubFamiliaEntity> lista = new List<SubFamiliaEntity>();

            string query = "SELECT RTRIM(Linea) Linea, RTRIM(Familia) Familia, RTRIM(SubFamilia) SubFamilia , RTRIM(DescripcionLocal) DescripcionLocal  FROM  WH_ClaseSubFamilia WHERE Linea=@idlinea AND Familia=@idfamilia AND Estado = 'A' ";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await connection.QueryAsync<SubFamiliaEntity>(query, new { idlinea, idfamilia });
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<MarcaEntity>> ListarMarca()
        {
            IEnumerable<MarcaEntity> lista = new List<MarcaEntity>();

            string query = "SELECT RTRIM(MarcaCodigo) MarcaCodigo, RTRIM(DescripcionLocal) DescripcionLocal, RTRIM(DescripcionIngles) DescripcionIngles  FROM  WH_Marcas WHERE ESTADO='A'";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await connection.QueryAsync<MarcaEntity>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<FormatoResponseRegistrarMaestroItem> RegistrarMaestroItem(DatosRequestMaestroItemModel dato,int idUsuario)
        {
            FormatoResponseRegistrarMaestroItem result = new FormatoResponseRegistrarMaestroItem();
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryFirstOrDefaultAsync<FormatoResponseRegistrarMaestroItem>("usp_Registrar_Item_sutura", new { NUMERO_PARTE=dato.codsut, CODIGO_FAMILIA=dato.familia, idUsuario }, commandType: CommandType.StoredProcedure);
            }
            return result;
        }

        public async Task<(List<FormatoListarMaestroItemModel>, int)> ListarMaestroItem(DatosListarMaestroItemPaginador datos)
        {
            (List<FormatoListarMaestroItemModel> ListaMaestroItems, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_Listar_Maestro_Items", datos, commandType: CommandType.StoredProcedure))
                {
                    result.ListaMaestroItems = result_db.Read<FormatoListarMaestroItemModel>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }

                connection.Dispose();
            }

            return result;
        }

        public async Task<IEnumerable<MaestroAlmacenEntity>> ListarMaestroAlmacen()
        {
            IEnumerable<MaestroAlmacenEntity> lista = new List<MaestroAlmacenEntity>();

            string query = "SELECT RTRIM(AlmacenCodigo) AlmacenCodigo, RTRIM(DescripcionLocal) DescripcionLocal, IIF (Estado='A','ACTIVO','INACTIVO') Estado  FROM WH_AlmacenMast  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await connection.QueryAsync<MaestroAlmacenEntity>(query);
                connection.Dispose();
            }

            return lista;
        }


        public async Task<bool> ValidacionPermisoAccesso( string permiso, int usuario)
        {
            bool result;

            using (var satelliteContext = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await satelliteContext.QuerySingleAsync<bool>("usp_validacion_PermisoAcceso",
                                    new { usuario, permiso }, commandType: CommandType.StoredProcedure);
                satelliteContext.Dispose();
            }

            return !result;
        }

        public async Task<DatosClienteDTO> ObtenerDatosCliente (int codigoCliente)
        {
            DatosClienteDTO cliente = new DatosClienteDTO();

            string query = "SELECT RTRIM(a.Documento) Documento, RTRIM(a.NombreCompleto) Nombre, IIF(c.Nacionalidad = 'E', 'EXTRANJERA', 'NACIONAL') Territorio, " +
                "RTRIM(b.DescripcionCorta) Pais, a.Estado FROM PersonaMast a INNER JOIN ClienteMast c ON a.Persona = c.Cliente " +
                "LEFT JOIN Pais b ON a.Nacionalidad = b.Pais WHERE Persona = @codigoCliente AND a.EsCliente = 'S'";

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                cliente = await connection.QueryFirstOrDefaultAsync<DatosClienteDTO>(query, new { codigoCliente });
            }

            return cliente;
        }

        public async Task ActualizarCorrelativoCodReclamo(int correlativo, int idConfiguracion, int id, string grupo)
        {
            string query = "UPDATE TBDConfiguracion SET ValorEntero1 = @correlativo WHERE IdConfiguracion = @idConfiguracion AND Id = @id AND Grupo = @grupo";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync(query, new { correlativo, idConfiguracion, id, grupo });
            }
        }

        public async Task<bool> ValidarExiteConfiguracionDetallePorId(int idConfiguracion, string estado = null)
        {
            bool validacion = false;

            string query = "SELECT 1 FROM TBDConfiguracion WHERE Id = idConfiguracion";

            if (!string.IsNullOrEmpty(estado))
                query = $"{query} AND Estado = @estado";

            query = QueryScript.ConsultaValidacion(query);

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                validacion = await connection.QueryFirstOrDefaultAsync<bool>(query, new { idConfiguracion, estado});
            }

            return validacion;
        }

        public async Task<IEnumerable<TipoDocumentoSsomaEntity>> TipoDocumentoSsoma()
        {
            IEnumerable<TipoDocumentoSsomaEntity> lista = new List<TipoDocumentoSsomaEntity>();

            string query = " SELECT idTipoDocumento, Descripcion ,Codigo , UsuarioCreacion , FechaModificacion , Estado FROM TBMSsomaTipoDocumento  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<TipoDocumentoSsomaEntity>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<UbicacionSsomaEntity>> UbicacionSsoma()
        {
            IEnumerable<UbicacionSsomaEntity> lista = new List<UbicacionSsomaEntity>();

            string query = " SELECT idUbicacionSsoma, Descripcion , UsuarioCreacion , FechaModificacion , Estado FROM TBMSsomaUbicacion  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<UbicacionSsomaEntity>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<ProteccionEntitySsoma>> ProteccionSsoma()
        {
            IEnumerable<ProteccionEntitySsoma> lista = new List<ProteccionEntitySsoma>();

            string query = " SELECT idProteccionSsoma, Descripcion , UsuarioCreacion , FechaModificacion , Estado FROM TBMSsomaProteccion  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<ProteccionEntitySsoma>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<EstadoEntitySsoma>> EstadoSsoma()
        {
            IEnumerable<EstadoEntitySsoma> lista = new List<EstadoEntitySsoma>();

            string query = " SELECT idEstadoSsoma, Descripcion , UsuarioCreacion , FechaModificacion , Estado FROM TBMSsomaEstado  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<EstadoEntitySsoma>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<AlmacenamientoSsomaEntity>> AlmacenamientoSsoma()
        {
            IEnumerable<AlmacenamientoSsomaEntity> lista = new List<AlmacenamientoSsomaEntity>();

            string query = " SELECT idSsomaAlmacenamiento, Descripcion , UsuarioCreacion , FechaModificacion , Estado FROM TBMSsomaAlmacenamiento  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<AlmacenamientoSsomaEntity>(query);
                connection.Dispose();
            }

            return lista;
        }

        public async Task<IEnumerable<ResponsableSsomaEntity>> ResponsableSsoma()
        {
            IEnumerable<ResponsableSsomaEntity> lista = new List<ResponsableSsomaEntity>();

            string query = " SELECT idSsomaResponsable, Descripcion , UsuarioCreacion , FechaModificacion , Estado FROM TBMSsomaResponsable  WHERE  Estado='A'; ";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await connection.QueryAsync<ResponsableSsomaEntity>(query);
                connection.Dispose();
            }

            return lista;
        }
    }
}
