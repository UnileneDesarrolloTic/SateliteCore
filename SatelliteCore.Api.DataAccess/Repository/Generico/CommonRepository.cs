using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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

        public async Task<List<FamiliaMP>> ListarFamiliaMP()
        {
            List<FamiliaMP> lista = new List<FamiliaMP>();
            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                lista = (List<FamiliaMP>)await connection.QueryAsync<FamiliaMP>("SELECT Codigo, Valor1 FROM TBDCatalogo WHERE Estado='A' and CatalogoID=6");
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



    }
}
