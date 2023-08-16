using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Encajado;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class EncajadoRespository : IEncajadoRespository
    {
        private readonly IAppConfig _appConfig;

        public EncajadoRespository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }


        public async Task<List<ListaOrdenesFabricaciónDTO>> ListaOrdenesFabricacion(string ordenFabricacion, string lote)
        {
            IEnumerable<ListaOrdenesFabricaciónDTO> lista = new List<ListaOrdenesFabricaciónDTO>();

            string query = "SELECT TOP 20 a.NumeroLote OrdenFabricacion, a.ReferenciaNumero Lote, b.NumeroDeParte Codsut, a.Item, b.DescripcionLocal Descripcion, a.CantidadProgramada, " +
                "a.Estado, ISNULL(c.Transferida, 0) Ingreso " +
                "FROM EP_ProgramacionLote a WITH(NOLOCK) INNER JOIN WH_ItemMast b WITH(NOLOCK) ON a.ITEM = b.Item " +
                "LEFT JOIN ( SELECT OrdenFabricacion, SUM(CantidadTransferida) Transferida FROM SatelliteCore.dbo.TBMEncaje WITH(NOLOCK) " +
                "GROUP BY OrdenFabricacion) c ON a.NumeroLote = c.OrdenFabricacion " +
                "WHERE a.CompaniaSocio = '01000000' AND a.Estado <> 'AN'";

            if (!string.IsNullOrWhiteSpace(ordenFabricacion))
                query = query + " AND a.NUMEROLOTE = @ordenFabricacion";

            if (!string.IsNullOrWhiteSpace(lote))
                query = query + " AND a.ReferenciaNumero = @lote";

            query = query + " ORDER BY c.OrdenFabricacion DESC, FechaProduccion DESC";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                lista = await context.QueryAsync<ListaOrdenesFabricaciónDTO>(query, new { ordenFabricacion, lote });
            }

            return lista.ToList();
        }

        public async Task<List<TransferenciaEncajadoDTO>> ListaTransferenciasEncaje(string ordenFabricacion)
        {
            IEnumerable<TransferenciaEncajadoDTO> lista = new List<TransferenciaEncajadoDTO>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await context.QueryAsync<TransferenciaEncajadoDTO>("usp_ListaTransferenciaEncaje", new { ordenFabricacion }, commandType: CommandType.StoredProcedure);
            }

            return lista.ToList();
        }

        public async Task<int> RegistarNuevaTrasnferencia(string ordenFabricacion, decimal cantidad, string usuario)
        {
            int registros = 0;

            string query = "INSERT INTO TBMEncaje (OrdenFabricacion, CantidadTransferida, UsuarioRegistro, FechaRegistro) VALUES(@ordenFabricacion, @cantidad, @usuario, GETDATE())";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                registros = await context.ExecuteAsync(query, new { ordenFabricacion, cantidad, usuario });
            }

            return registros;
        }

        public async Task<(decimal, decimal, List<AsignacionEncajadoDTO>)> ListraAsignacionesEncajePorEtapa(int idEncaje, int etapa)
        {
            IEnumerable<AsignacionEncajadoDTO> lista = new List<AsignacionEncajadoDTO>();
            decimal totalAnterior = 0;
            decimal totalActual = 0;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await context.QueryMultipleAsync("usp_listaAsignacionesPorEtapa", new { idEncaje, etapa }, commandType: CommandType.StoredProcedure))
                {
                    totalAnterior = result_db.Read<decimal>().FirstOrDefault();
                    totalActual = result_db.Read<decimal>().FirstOrDefault();
                    lista = result_db.Read<AsignacionEncajadoDTO>().ToList();
                }
            }

            return (totalAnterior, totalActual, lista.ToList());
        }

        public async Task<int> RegistrarAsignacionEncaje(DatosRegistrarAsignacionDTO asignacion)
        {
            int registros = 0;
            int codigoEmpleado = 0;
            string query = "INSERT INTO TBDEncaje (IdEncaje, Etapa, Usuario, Cantidad, Estado, Fecha, UsuarioRegistro, FechaRegistro) " +
                "VALUES(@idEncaje, @etapa, @empleado, @cantidad, 'A', @fecha, @usuarioRegistro, GETDATE())";

            string queryValidacion = "SELECT Empleado FROM EmpleadoMast WITH(NOLOCK) WHERE Empleado = @empleado AND Estado = 'A'";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                codigoEmpleado = await context.QueryFirstOrDefaultAsync<int>(queryValidacion, new { asignacion.Empleado });
            }

            if (codigoEmpleado != asignacion.Empleado)
                return -1;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                registros = await context.ExecuteAsync(query, asignacion);
            }

            return registros;
        }

        public async Task ActualizaEstadoAsignacion(int id, string estado, string usuario)
        {
            string query = "UPDATE TBDEncaje SET Estado = @estado, UsuarioCompletado = @usuario, FechaCompletado = GETDATE() WHERE Id = @id";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(query, new { id, estado, usuario });
            }
        }


        public async Task<List<DatosReporteEncajadoDTO>> DatosReporteAsignacion(DateTime fechaInicio, DateTime fechaFin)
        {
            IEnumerable<DatosReporteEncajadoDTO> lista = new List<DatosReporteEncajadoDTO>();
            string query = "SELECT a.Id, a.OrdenFabricacion, a.CantidadTransferida, a.FechaRegistro, " +
                "CASE b.Etapa WHEN 1 THEN 'ARMADO CAJA' WHEN 2 THEN 'ENCAJADO' WHEN 3 THEN 'SELLADO' WHEN 4 THEN 'EMBALADO' ELSE '' END Etapa, " +
                "RTRIM(c.NombreCompleto) UsuarioAsignado, b.Cantidad CantidadAsignada, b.FechaRegistro FechaAsignada, " +
                "CASE b.Estado WHEN 'C' THEN 'Completado' WHEN 'A' THEN 'Asignado' ELSE '' END Estado " +
                "FROM TBMEncaje a WITH(NOLOCK) INNER JOIN TBDEncaje b WITH(NOLOCK) ON a.Id = b.IdEncaje " +
                "INNER JOIN PROD_UNILENE2..PersonaMast c WITH(NOLOCK) ON b.Usuario = c.Persona AND c.EsEmpleado = 'S' " +
                "WHERE b.Estado IN ('A','C')";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                lista = await context.QueryAsync<DatosReporteEncajadoDTO>(query, new { fechaInicio, fechaFin });
            }

            return lista.ToList();
        }

    }
}
