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

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class UsuarioRepository : IUsuarioRepository
    {
        private readonly IAppConfig _appConfig;

        public UsuarioRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<UsuarioEntity> ObtenerUsuario(DatoUsuarioBasico datos)
        {
            UsuarioEntity usuario = new UsuarioEntity ();

            string script = "SELECT IDUsuario, Nombre, ApellidoPaterno, ApellidoMaterno, TipoDocumento, " +
                    "NroDocumento, Sexo, Pais, Correo, Estado, FechaNacimiento, Celular FROM TBMUsuario " +
                    "WHERE IDUsuario = @Id";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                usuario = await connection.QueryFirstOrDefaultAsync<UsuarioEntity>(script, new { Id = datos.IdUsuario });
            }

            return usuario;
        }

        public async Task<AuthResponse> ObtenerUsuarioLogin(AuthRequestModel datosUsuario)
        {
           
            AuthResponse usuario = new AuthResponse();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "SELECT a.CodUsuario, a.Usuario, RTRIM(b.Nombres) Nombres, RTRIM(b.ApellidoPaterno) ApellidoPaterno, a.Correo FROM TBMUsuario a " +
                    " INNER JOIN PROD_UNILENE2.dbo.PersonaMast b ON a.CodUsuario = b.Persona " +
                    "WHERE a.Estado = 'A' AND b.Estado = 'A' AND Usuario = @Usuario AND Clave = @Clave";
                usuario = await connection.QueryFirstOrDefaultAsync<AuthResponse>(sql, datosUsuario);
            }
            return usuario;

        }

        public async Task<int> CambiarClave(ActualizarClaveModel datos)
        {
            int result = 0;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                string sql = "UPDATE TBMUsuario SET Clave = @clave, FlagCambioClave = @exigirCambioClave" +
                    "WHERE IDUsuario = @idUsuario AND NroDocumento = @nroDocumento AND Estado = 'A'";

                result = await connection.ExecuteAsync(sql, datos);
                connection.Dispose();
            }

            return result;
        }

        public async Task<(List<UsuarioEntity>, int)> ListarUsuarios(DatosListarUsuarioPaginado datos)
        {
            (List<UsuarioEntity> ListaUsuarios, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_listar_Usuario", datos, commandType: CommandType.StoredProcedure))
                {
                    result.ListaUsuarios = result_db.Read<UsuarioEntity>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }

                connection.Dispose();
            }

            return result;
        }

        public async Task<DatosFormatoAsignacionPersonalLaboralModel> ListarAsignacionPersonal()
        {
            //IEnumerable<FormatoDeAsignacionPersonalLaboralModel> result = new List<FormatoDeAsignacionPersonalLaboralModel>();

            DatosFormatoAsignacionPersonalLaboralModel result = new DatosFormatoAsignacionPersonalLaboralModel();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_unilene_pr_listar_horario_personal", commandType: CommandType.StoredProcedure))
                {
                    result.PersonalLaboral = result_db.Read<FormatoDeAsignacionPersonalLaboralModel>().ToList();
                    result.ContarArea = result_db.Read<DatosFormatoContarAreaModel>().ToList();
                }
            }

            return result;
        }

        public async Task<List<AreaPersonalLaboralEntity>> ListarAreaPersonaLaboral()
        {
            List<AreaPersonalLaboralEntity> result = new List<AreaPersonalLaboralEntity>();
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = (List<AreaPersonalLaboralEntity>)await connection.QueryAsync<AreaPersonalLaboralEntity>("usp_ListarAreas", commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<int> RegistrarPersonaLaboralMasiva(DatosFormatoAsignacionPersonalModel dato,int idUsuario)
        {
            int result=1;
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {   
                foreach (PersonaLaboral valor in dato.ListaPersona)
                {
                    await connection.ExecuteAsync("sp_RegistrarPersonalArea", new { dato.IdArea, valor.idPersona ,idUsuario }, commandType: CommandType.StoredProcedure);
                }
                
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel>> FiltrarAreaPersona(int idArea,string NombrePersona)
        {
            IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel> result = new List<DatosFormatoFiltrarTrabajadorAreaModel>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatoFiltrarTrabajadorAreaModel>("usp_Filtrar_Empleado_por_Area", new { idArea, NombrePersona }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<int> LiberalPersona(int IdAsignacion)
        {
            int result = 1;
            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.ExecuteAsync("sp_LiberarPersonalArea", new { IdAsignacion }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoPersonaAsignacionExportModel>> ExportarExcelPersonaAsignacion(string FechaInicio, string FechaFinal)
        {
            IEnumerable<DatosFormatoPersonaAsignacionExportModel> result = new List<DatosFormatoPersonaAsignacionExportModel>();

            string query = "SELECT a.IdAsignacion, RTRIM(b.Busqueda) NombreCompleto ,c.Descripcion NombreArea,a.FechaAsignacion ,a.FechaReAsignacion , IIF(a.Estado='A','Activo','Inactivo') Estado " +
                          "FROM TBMAsignacionArea a " +
                          "INNER JOIN PROD_UNILENE2..PersonaMast b on b.Persona = a.IdEmpleado " +
                          "INNER JOIN TBMAreasProduccion c ON a.IdArea = c.IdArea " +
                          "WHERE(CONVERT(varchar, a.FechaAsignacion, 23) >= @FechaInicio AND CONVERT(varchar, a.FechaAsignacion, 23) <= @FechaFinal)";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatoPersonaAsignacionExportModel>(query, new { FechaInicio, FechaFinal });
            }

            return result;
        }



        public async Task<AreaPersonalLaboralEntity> RegistrarEditarArea(int IdArea, string Descripcion)
        {
            string sql = "";
            AreaPersonalLaboralEntity result =  new AreaPersonalLaboralEntity();

            if (IdArea == 0)
                sql = "INSERT INTO TBMAreasProduccion (Descripcion,Estado) VALUES (@Descripcion,'A');" +
                      "SELECT IdArea,Descripcion FROM TBMAreasProduccion WHERE IdArea=@@IDENTITY";
            else
                sql = "UPDATE TBMAreasProduccion SET Descripcion=@Descripcion WHERE IdArea=@IdArea;" +
                      "SELECT IdArea,Descripcion FROM TBMAreasProduccion WHERE IdArea=@IdArea";

                using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
                {
                    result=await connection.QueryFirstOrDefaultAsync<AreaPersonalLaboralEntity>(sql, new { IdArea, Descripcion });
                }

            return result;
        }

        public async Task<int> EliminarAreaProduccion(int IdArea)
        {
            string sql = "UPDATE TBMAreasProduccion SET Estado='I' WHERE IdArea=@IdArea; " +
                         "UPDATE TBMAsignacionArea SET TBMAsignacionArea.Estado='I' FROM TBMAsignacionArea " +
                         "INNER JOIN PROD_UNILENE2..PersonaMast b on b.Persona = TBMAsignacionArea.IdEmpleado " +
                         "INNER JOIN TBMAreasProduccion c ON TBMAsignacionArea.IdArea = c.IdArea " +
                         "WHERE CONVERT(varchar, TBMAsignacionArea.FechaAsignacion,103)= CONVERT(varchar, SYSDATETIME(), 103) AND c.IdArea =@IdArea";
            int result = 0;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await connection.QueryAsync(sql, new { IdArea });
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoListarPersonaTecnica>> ListarPersonaTecnico()
        {
            IEnumerable<DatosFormatoListarPersonaTecnica> result = new List<DatosFormatoListarPersonaTecnica>();

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatoListarPersonaTecnica>("usp_Listar_persona_unilene_tecnico", commandType: CommandType.StoredProcedure);
            }

            return result;
        }



    }
}
