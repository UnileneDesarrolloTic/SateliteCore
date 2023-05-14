using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Report.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.RRHH;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.RRHH;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class RRHHRepository : IRRHHRepository
    {
        private readonly IAppConfig _appConfig;

        public RRHHRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ReporteAsistenciaDTO>> ListarAsistencia(DateTime fecha, int usuarioToken)
        {
            using SqlConnection context = new SqlConnection(_appConfig.contextSpring);
            IEnumerable<ReporteAsistenciaDTO> listaAsistencia = await context.QueryAsync<ReporteAsistenciaDTO>("usp_ObtenerReporteDiarioAsistencia_Satelite", new { fecha, usuario = usuarioToken }, commandType: CommandType.StoredProcedure);

            return listaAsistencia;
        }

        public async Task<int> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera data, string usuario)
        {
            int result = 0;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                int IdCabecera = await connection.QueryFirstOrDefaultAsync<int>("usp_InsertarHorasExtrasCabecera", new { data.IdCodigo, data.Area, data.Persona, data.Justificacion, data.FechaRegistro, usuario, data.Estado }, commandType: CommandType.StoredProcedure);

                foreach (DatosEstructuraHorasExtrasDetalle persona in data.ListaPersona)
                {
                    await connection.ExecuteAsync("INSERT INTO TBMHoraExtrasDetalle (IdCabecera,IdPersona,FechaCreacion,cant_horas) VALUES (@IdCabecera,@persona,GETDATE(),@horasextras);", new { data.IdCodigo, IdCabecera, persona.persona, persona.horasextras });
                }

                connection.Dispose();
            }
            return result;
        }

        public async Task<IEnumerable<DatosFormatoListarHorasExtrasPersona>> ListarHoraExtrasPersona(DatosFormatoFiltraHoraExtras data)
        {
            IEnumerable<DatosFormatoListarHorasExtrasPersona> result = new List<DatosFormatoListarHorasExtrasPersona>();

            string query = "SELECT a.IdCabecera, a.IdArea,b.Descripcion NombreArea,a.FechaRegistro," +
                                "(CASE a.TipoPersona WHEN 'E' THEN 'EMPLEADO' ELSE 'OBRERO' END) TipoPersona," +
                                "(CASE a.Estado WHEN 'AP' THEN 'APROBADO' WHEN 'GE' THEN 'GENERADO' WHEN 'PR' THEN 'PROCESADO' ELSE 'ANULADO' END) Estado, " +
                                "a.Justificacion, ISNULL(a.UsuarioAprobacion,'-----------')  UsuarioAprobacion, ISNULL(a.Periodo,'') Periodo," +
                                "a.FechaAprobacion " +
                            "FROM  TBMHoraExtrasCabecera a " +
                            "INNER JOIN TBMAreasProduccion b ON b.IdArea = a.IdArea WHERE 1=1";

            if (data.Estado != "TD")
                query = $"{query} AND a.Estado = @Estado";

            if (data.FechaInicio != null && data.FechaFin != null)
                query = $"{query} AND a.FechaRegistro BETWEEN @FechaInicio AND @FechaFin";
            else
                if (data.Periodo != null)
                query = $"{query} AND a.Periodo = @Periodo";

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await connection.QueryAsync<DatosFormatoListarHorasExtrasPersona>(query, new { data.Estado, data.FechaInicio, data.FechaFin, data.Periodo });
            }

            return result;
        }

        public async Task ProcesarHorasExtrasPlanilla(string periodo)
        {
            using SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB);
            _ = await context.ExecuteAsync("usp_RRHH_ProcesarHorasExtrasPlanilla", new { periodo }, commandType: CommandType.StoredProcedure);
        }

        public async Task<(DatosFormatoHorasExtrasCabeceraModel, List<DatosFormatoHorasExtrasDetalle>)> BuscarInformacionHorasExtrasPersona(int Cabecera)
        {
            (DatosFormatoHorasExtrasCabeceraModel cabecera, List<DatosFormatoHorasExtrasDetalle> detalle) HorasExtrasPersona;

            string sql = "SELECT a.IdCabecera,a.IdArea,b.Descripcion,a.Justificacion,a.FechaRegistro,a.TipoPersona, a.Estado, ISNULL(a.Periodo,'') Periodo  FROM TBMHoraExtrasCabecera a  " +
                         "INNER JOIN  TBMAreasProduccion b ON a.IdArea = b.IdArea WHERE a.IdCabecera = @Cabecera ; " +
                         "SELECT c.IdDetalle,c.IdCabecera,c.IdPersona,RTRIM(d.Busqueda) NombrePersona, RTRIM(ISNULL(d.Documento,d.DocumentoIdentidad)) Documento , c.cant_horas horasextras  FROM TBMHoraExtrasDetalle c  " +
                         "INNER JOIN PROD_UNILENE2..PersonaMast d ON c.IdPersona = d.Persona  WHERE c.IdCabecera = @Cabecera ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync(sql, new { Cabecera });
                HorasExtrasPersona.cabecera = multi.Read<DatosFormatoHorasExtrasCabeceraModel>().FirstOrDefault();
                HorasExtrasPersona.detalle = multi.Read<DatosFormatoHorasExtrasDetalle>().ToList();
            }

            return HorasExtrasPersona;
        }

        public async Task<List<HorasExtraExportDTO>> ListarHorasExtrasGeneradas(string periodo)
        {
            IEnumerable<HorasExtraExportDTO> lista = new List<HorasExtraExportDTO>();
            string query = "SELECT a.FechaRegistro, b.IdPersona, b.cant_horas, " +
                    "IIF(cant_horas <= 2, cant_horas, 2) H25, IIF(cant_horas > 2, cant_horas - 2, 0) H35, " +
                    "c.SueldoActualLocal, ROUND((c.SueldoActualLocal / 30 / 8), 2) SueldoHora, " +
                    "RTRIM(d.NombreCompleto) NombreCompleto, " +
                    "ROUND((IIF(cant_horas > 2, cant_horas - 2, 0) * ((c.SueldoActualLocal / 30 / 8) * 1.25)), 2) 'S25', " +
                    "ROUND((IIF(cant_horas <= 2, cant_horas, 2) * ((c.SueldoActualLocal / 30 / 8) * 1.35)), 2) 'S35' " +
                "FROM TBMHoraExtrasCabecera a " +
                    "INNER JOIN TBMHoraExtrasDetalle b ON a.idcabecera = b.idcabecera " +
                    "INNER JOIN PROD_UNILENE2..EmpleadoMast c ON b.IdPersona = c.Empleado " +
                    "INNER JOIN  PROD_UNILENE2..PersonaMast d ON d.Persona = c.Empleado " +
                "WHERE a.periodo = @periodo and a.estado = 'PR' " +
                "ORDER BY d.nombrecompleto, a.fecharegistro";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await context.QueryAsync<HorasExtraExportDTO>(query, new { periodo });
            }

            return lista.ToList();
        }

        public async Task<AutorizacionSobretiempoPersonaDTO> ListarHorasExtraExportas(string periodo)
        {
            AutorizacionSobretiempoPersonaDTO result = new AutorizacionSobretiempoPersonaDTO();

            string query = "SELECT IdPersona, RTRIM(c.NombreCompleto) Nombres, f.Descripcion Area, " +
                   "a.FechaRegistro,CONVERT(VARCHAR(5), a.FechaRegistro, 8) HoraInicio," +
                   "CONVERT(VARCHAR(5), DATEADD(MINUTE, ((b.Cant_horas - FLOOR(b.Cant_horas)) * 60), DATEADD(HOUR,cant_horas, FechaRegistro)), 8) HoraFin," +
                   "CONVERT(VARCHAR, FLOOR(b.Cant_horas)) + ':'+ IIF( FLOOR((b.Cant_horas - FLOOR(b.Cant_horas)) * 60) < 10, '0' + CONVERT(VARCHAR, FLOOR((b.Cant_horas - FLOOR(b.Cant_horas)) * 60)), CONVERT(VARCHAR, FLOOR((b.Cant_horas - FLOOR(b.Cant_horas)) * 60)) ) Cant_horas, " +
                   "RTRIM(UPPER(g.LocalName)) CentroCosto, d.Descripcion SubArea, ROW_NUMBER() OVER(PARTITION BY IdPersona ORDER BY IdPersona) Contador " +
               "INTO #temp_HorasExtras " +
               "FROM TBMHoraExtrasCabecera a " +
                   "INNER JOIN TBMHoraExtrasDetalle b ON a.IdCabecera = b.IdCabecera " +
                   "INNER JOIN PROD_UNILENE2.dbo.PersonaMast c ON b.IdPersona = c.Persona " +
                   "INNER JOIN TBMAreasProduccion d ON a.IdArea = d.IdArea " +
                   "INNER JOIN PROD_UNILENE2.dbo.EmpleadoMast e ON c.Persona = e.Empleado " +
                   "INNER JOIN PROD_UNILENE2.dbo.HR_PuestoEmpresa f ON e.CodigoCargo = f.CodigoPuesto " +
                   "LEFT JOIN  PROD_UNILENE2.dbo.AC_CostCenterMst g ON e.CentroCostos = g.CostCenter " +
               "WHERE Periodo = @Periodo AND a.Estado = 'PR'; " +

               "SELECT IdPersona, Nombres, Area, CentroCosto, SubArea FROM #temp_HorasExtras WHERE Contador = 1; " +

               "SELECT IdPersona, FechaRegistro, HoraInicio, HoraFin, Cant_horas FROM #temp_HorasExtras " +

               "DROP TABLE #temp_HorasExtras";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync(query, new { periodo });
                result.Cabecera = multi.Read<AutoSobretiempoPersonaCabeceraDTO>().ToList();
                result.Detalle = multi.Read<AutoSobretiempoPersonaDetalleDTO>().ToList();
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoReporteComisionVendedor>> ReporteComisionVendedor(string periodo)
        {
            IEnumerable<DatosFormatoReporteComisionVendedor> result = new List<DatosFormatoReporteComisionVendedor>();

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                result = await connection.QueryAsync<DatosFormatoReporteComisionVendedor>("usp_reporte_comisiones_vendedor_satelite", new { periodo }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

    }
}
