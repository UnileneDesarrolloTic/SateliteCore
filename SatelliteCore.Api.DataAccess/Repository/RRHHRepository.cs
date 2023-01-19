﻿using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.RRHH;
using SatelliteCore.Api.Models.Response;
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
            IEnumerable<ReporteAsistenciaDTO> listaAsistencia = new List<ReporteAsistenciaDTO>();

            using SqlConnection context = new SqlConnection(_appConfig.contextSpring);
            listaAsistencia = await context.QueryAsync<ReporteAsistenciaDTO>("usp_ObtenerReporteDiarioAsistencia_Satelite", new { fecha, usuario = usuarioToken }, commandType: CommandType.StoredProcedure);

            return listaAsistencia;
        }

        public async Task<int> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera data, string usuario)
        {
            int result = 0;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                int IdCabecera = await connection.QueryFirstOrDefaultAsync<int>("usp_InsertarHorasExtrasCabecera", new { data.idCodigo, data.Area, data.Persona, data.Justificacion, data.FechaRegistro, usuario, data.Estado }, commandType: CommandType.StoredProcedure);

                foreach (DatosEstructuraHorasExtrasDetalle persona in data.ListaPersona)
                {
                    await connection.ExecuteAsync("INSERT INTO TBMHoraExtrasDetalle (IdCabecera,IdPersona,FechaCreacion,cant_horas) VALUES (@IdCabecera,@persona,GETDATE(),@horasextras);", new { data.idCodigo, IdCabecera, persona.persona, persona.horasextras });
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
                                "(CASE a.Estado WHEN 'AP' THEN 'APROBADO' WHEN 'GE' THEN 'GENERADO' WHEN 'PR' THEN 'PROCESADO   ' ELSE 'ANULADO' END) Estado, " +
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
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                _ = await context.ExecuteAsync("usp_RRHH_ProcesarHorasExtrasPlanilla", new { periodo }, commandType: CommandType.StoredProcedure);
            }
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
    }
}
