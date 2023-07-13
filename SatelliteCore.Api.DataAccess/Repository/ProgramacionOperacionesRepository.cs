﻿using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.ProgramacionOperaciones;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.ProgramacionOperaciones;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class ProgramacionOperacionesRepository :  IProgramacionOperacionesRepository
    {
        private readonly IAppConfig _appConfig;

        public ProgramacionOperacionesRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<DatosFormatoAgrupadores>> ObtenerAgrupadores(string gerencia)
        {
            IEnumerable<DatosFormatoAgrupadores> result = new List<DatosFormatoAgrupadores>();

            string sql = ";WITH Temp_Agrupador AS(" +
                "SELECT DISTINCT DPRO_ORDE_AGRUPADOR Agrupador, DPRO_CODI_AGRUPADOR_2 idAgrupador, (case when DPRO_CODI_AGRUPADOR_2 = 1 then 'Suturas' " +
                "WHEN DPRO_CODI_AGRUPADOR_2 in (18,19 ) THEN 'Plasticos' WHEN DPRO_DESC_SUB_AGRUPADOR in ('CANULA YANKAUER','CLAMPS UMBILICAL','TUBOS DE ASPIRACION') or DPRO_DESC_FAMILIA in('MANGA POLIETILENO') then 'Plasticos' " +
                "ELSE 'Demas agrupadores' END) as Gerencia FROM DIM_PRODUCTO  WHERE DPRO_CODI_LINEA = 'P'" +
                ")" +
                "SELECT Agrupador, idAgrupador, Gerencia FROM Temp_Agrupador WHERE ";
            if (gerencia == "Suturas")
                sql = $"{sql} Gerencia = 'Suturas'";
            else if (gerencia == "Plasticos")
                sql = $"{sql} Gerencia = 'Plasticos'";
            else
                sql = $"{sql} Gerencia = 'Demas agrupadores'";
            
            using (SqlConnection context = new SqlConnection(_appConfig.ContextDMVentas))
            {
                result = await context.QueryAsync<DatosFormatoAgrupadores> (sql);
            }

            return result;
        }


        public async Task<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>> ObtenerProgramacionOrdenFabricacion(DatosFormatoProgramacionOperaciones dato, string unionAgrupador)
        {
            IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion> result = new List<DatosFormatoProgramacionOperacionesOrdenFabricacion>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<DatosFormatoProgramacionOperacionesOrdenFabricacion>("usp_satelite_lista_ProgramacionOperaciones_OrdenFabricacion", new { dato.gerencia, agrupador= unionAgrupador, dato.fechaInicio, dato.fechaFinal, dato.lote, dato.ordenFabricacion, dato.venta, dato.estado }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<string> ActualizarFechaProgramada(DatosFormatoRegistrarFechaProgramacion dato, string usuario)
        {
            string result = "";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryFirstOrDefaultAsync<string>("sp_Satelite_Registro_FechaProgramacion", new { dato.ordenFabricacion, dato.fechaInicio, dato.fechaEntrega, dato.comentarioInicio, dato.comentarioEntrega, usuario }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

        public async Task<IEnumerable<DatosFormatoListadoFechaProgramadas>> ObtenerTipoFechaOrdenFabricacion(string ordenFabricacion, string tipoFecha)
        {
            IEnumerable<DatosFormatoListadoFechaProgramadas> result = new List<DatosFormatoListadoFechaProgramadas>();
            string sql = "SELECT OrdenFabricacion, Tipo, Fecha, Comentario, UsuarioCreacion, FechaRegistro FROM TBDFechaProgramacionHistorica WHERE OrdenFabricacion = @ordenFabricacion  AND Tipo = @tipoFecha ORDER BY FechaRegistro DESC ";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                result = await context.QueryAsync<DatosFormatoListadoFechaProgramadas>(sql, new { ordenFabricacion, tipoFecha });
            }

            return result;
        }
        public async Task<string> RegistrarDivisionProgramacion(DatosFormatoDividirRegistroProgramacion dato, string usuario)
        {
            string result = "";
            string sql = "DELETE FROM SatelliteCore..TBDDividirProgramacion WHERE OrdenFabricacion = @ordenFabricacion";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                await context.ExecuteAsync(sql, new { dato.ordenFabricacion });

                foreach (DatosFormatoDivisionProgramacion division in dato.divisionProgramacion)
                    await context.ExecuteAsync("usp_Satelite_Registro_Dividir_Programacion", new { dato.ordenFabricacion, division.correlactivo, division.cantidad, usuario }, commandType: CommandType.StoredProcedure);
            }

            return result;
        }

    }
}
