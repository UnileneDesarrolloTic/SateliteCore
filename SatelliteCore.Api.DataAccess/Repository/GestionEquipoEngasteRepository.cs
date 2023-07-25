using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Response.GestioEquipoEngaste;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class GestionEquipoEngasteRepository : IGestionEquipoEngasteRepository 
    {
        private readonly IAppConfig _appConfig;

        public GestionEquipoEngasteRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<DatosFormatoEmpleado>> ObtenerEmpleado()
        {
            IEnumerable<DatosFormatoEmpleado> result = new List<DatosFormatoEmpleado>();

            string sql = "SELECT b.persona, RTRIM(ISNULL(b.NombreCompleto, b.busqueda)) NombreCompleto FROM EmpleadoMast a INNER JOIN PersonaMast b ON a.Empleado = b.Persona WHERE a.Estado = 'A' AND a.Empleado NOT IN ('-1') ;";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatoEmpleado>(sql);
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoListadoDadoEngaste>> ObtenerListadoDados()
        {
            IEnumerable<DatosFormatoListadoDadoEngaste> result = new List<DatosFormatoListadoDadoEngaste>();

            string sql = "SELECT idDadoEngasteCodificacion Id, Codigo, Dado, Alambre, Tipo,  CAST(0 AS BIT) Seleccionar  FROM TBDDadoEngasteCodificacion WHERE Estado = 'A';";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoListadoDadoEngaste>(sql);
            }

            return result;
        }
        public async Task<IEnumerable<DatosFormatoListarEquipoEngaste>> ListarEquipoEngaste(DatosFormularioFiltroEquipo dato)
        {
            IEnumerable<DatosFormatoListarEquipoEngaste> result = new List<DatosFormatoListarEquipoEngaste>();

            string sql = "SELECT DISTINCT b.IdEquipo, b.Descripcion, b.Tipo, b.IdPersona, RTRIM(a.NombreCompleto) NombreCompleto, b.Estado, b.FechaCreacion " +
                         "FROM PersonaMast a INNER JOIN SatelliteCore..TBMGestionEquipoEngaste b ON a.Persona = b.IdPersona " +
                         "LEFT JOIN SatelliteCore..TBDDetalleGestionEquiposEngaste c ON b.IdEquipo = c.IdEquipo " +
                         "WHERE 1 = 1 " +
                         (dato.persona == 0 ? "" : " AND b.IdPersona = @persona ") + (string.IsNullOrEmpty(dato.tipo) ? "" : " AND b.Tipo = @tipo ") + (dato.codigo == 0 ? "" : " AND  c.IdDadoEngasteCodificacion = @codigo ") ;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatoListarEquipoEngaste>(sql, new { dato.persona, dato.tipo, dato.codigo });
            }

            return result;
        }

        public async Task<DatosFormatoInformacionEquipoEngaste> ObtenerInformacionEquipo(string idEquipo)
        {
            DatosFormatoInformacionEquipoEngaste result = new DatosFormatoInformacionEquipoEngaste();

            using (SqlConnection DMVentasContext = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await DMVentasContext.QueryMultipleAsync("usp_informacion_equipo_engaste", new { idEquipo }, commandType: CommandType.StoredProcedure);
                result.cabecera = multi.Read<DatosFormatoCabeceraInformacionEquipo>().FirstOrDefault();
                result.detalle = multi.Read<DatosFormatoDetalleInformacionEquipo>().ToList();

            }

            return result;
        }



        public async Task<string> RegistrarEquipoEngastado(DatosFormatoRegistroEquipoEngastado dato)
        {
            int id ;
            string resultado = "";
            string sqlDetalle = "INSERT INTO TBDDetalleGestionEquiposEngaste (IdEquipo, IdDadoEngasteCodificacion) VALUES (@id, @idDado);";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                id = await context.QueryFirstOrDefaultAsync<int>("usp_Registro_Equipo_Engaste", new { dato.idEquipo, dato.nombre, dato.idpersona, dato.Tipo, dato.estado }, commandType: CommandType.StoredProcedure);

                foreach (int idDado in dato.detalle)
                    await context.QueryFirstOrDefaultAsync<int>(sqlDetalle, new { id, idDado });
                
            }


            return resultado;
        }

    }
}
