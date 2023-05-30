using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class OrdenServicioRepository : IOrdenServicioRepository
    {
        private readonly IAppConfig _appConfig;

        public OrdenServicioRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ListarOrdenServicioResponseDTO>> ListarOrdenServicio(DateTime fechaInicio, DateTime fechaFin)
        {
            IEnumerable<ListarOrdenServicioResponseDTO> ordenes = new List<ListarOrdenServicioResponseDTO>();

            string query = "SELECT a.Id, a.FECHA, a.FechaSalida, a.NUMERO_OS OrdenServicio, b.id IdTransportista, b.Descripcion Transportista, a.Usuario, a.Estado " +
                "FROM TLOG_PLAN_ORDEN_SERVICIO_H a LEFT JOIN TLO_TRANSPORTISTA b ON a.TRANSPORTISTA = b.ID " +
                "WHERE a.Estado = 'A' AND CAST(a.Fecha AS DATE) BETWEEN @fechaInicio AND @fechaFin ORDER BY a.Id DESC";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                ordenes = await context.QueryAsync<ListarOrdenServicioResponseDTO>(query, new { fechaInicio, fechaFin });
            }

            return ordenes;
        }

        public async Task<IEnumerable<ListaTransportistaComboxResponse>> ListarTransportistaCombox()
        {
            IEnumerable<ListaTransportistaComboxResponse> transportistas = new List<ListaTransportistaComboxResponse>();

            string query = "SELECT Id, Descripcion FROM UNILENE_REPORTEADOR..TLO_TRANSPORTISTA WHERE ESTADO='A'";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                transportistas = await context.QueryAsync<ListaTransportistaComboxResponse>(query);
            }

            return transportistas;
        }

        public async Task<IEnumerable<DetalleOrdenServicioResponse>> ListaDetalleOrdenServicio(int codigoOrdenServicio)
        {
            IEnumerable<DetalleOrdenServicioResponse> detalle = new List<DetalleOrdenServicioResponse>();

            string query = "SELECT Id, ISNULL(Serie, '') + '-' + ISNULL(Numero_Documento, '') AS Guia, Fecha_Documento Fecha, Cliente_Descripcion Cliente, " +
                "Direccion_Destino Direccion, Destino Departamento, Factura_Numero Factura, Peso, Bultos, Comentarios Comentario " +
                "FROM TLOG_PLAN_ORDEN_SERVICIO_D WHERE Estado = 'A' AND ID_H = @codigoOrdenServicio";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                detalle = await context.QueryAsync<DetalleOrdenServicioResponse>(query, new { codigoOrdenServicio });
            }

            return detalle;
        }

        public async Task<IEnumerable<OrdenServicioGuiaRemisionResponse>> ListaGuiaRemision(DateTime fechaInicio, DateTime  fechaFin)
        {
            IEnumerable<OrdenServicioGuiaRemisionResponse> guias = new List<OrdenServicioGuiaRemisionResponse>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                guias = await context.QueryAsync<OrdenServicioGuiaRemisionResponse>("usp_OrdenesServicio_listaGuias", new { fechaInicio, fechaFin }, commandType: CommandType.StoredProcedure);
            }

            return guias;
        }

        public async Task Modificar_Peso_Bultos(List<OrdenServicioDetalle> ordenes)
        {
            string query = "UPDATE TLOG_PLAN_ORDEN_SERVICIO_D SET Peso = @peso, Bultos = @bultos WHERE Id = @id";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(query, ordenes);
            } 
        }

        public async Task ModificarTransportista(int id, int idTransportista)
        {
            string query = "UPDATE TLOG_PLAN_ORDEN_SERVICIO_H SET TRANSPORTISTA = @idTransportista WHERE ID = @id";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(query, new { id, idTransportista });
            }
        }

        public async Task RegistrarGuias_OrdenServicio (List<RegistrarGuia_OrdenServicioDTO> datosGuia)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                await context.ExecuteAsync("usp_OrdenServicio_RegistrarOS_Detalle", datosGuia, commandType: CommandType.StoredProcedure);
            }
        }

        public async Task RegistrarObjetosExtrasEnvio_OS (List<OrdenServicioDetalle> extras)
        {
            string query = "INSERT INTO TLOG_PLAN_ORDEN_SERVICIO_D (ID_H, NUMERO_DOCUMENTO, FECHA_DOCUMENTO, CLIENTE_DESCRIPCION, " +
                "DEPARTAMENTO_CLIENTE, FACTURA_NUMERO, DESTINO, DIRECCION_DESTINO, PESO, BULTOS, USUARIO, ESTADO ) " +
                "VALUES(@Cabecera   , @Guia, @Fecha, @Cliente, @Departamento, @Comercial, @Departamento, @Direccion, @Peso, @Bultos, @Usuario, 'A')";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(query, extras);
            }
        }

        public async Task EliminarDetalleOrdenServicio (int id)
        {
            string query = "UPDATE TLOG_PLAN_ORDEN_SERVICIO_D SET ESTADO = 'I' WHERE ID = @id";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(query, new { id } );
            }
        }

        public async Task EditarGuiaRemision(EditarGuiaOS_DTO datosGuia)
        {
            string query = "UPDATE TLOG_PLAN_ORDEN_SERVICIO_D SET FECHA_RETORNO = @fecha, DEPARTAMENTO_CLIENTE = @departamento, PESO = @peso, " +
                "BULTOS = @bultos, CLIENTE_DESCRIPCION = @cliente, DIRECCION_DESTINO = @direccion, COMENTARIOS = @comentario WHERE ID = @id and ID_H = @cabecera";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(query, datosGuia);
            }
        }

        public async Task GuardarTransportista(DatosTransportistaDTO datosTransportista)
        {
            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync("SP_TLO_INSERTAR_TLO_TRANSPORTISTA", datosTransportista, commandType: CommandType.StoredProcedure);
            }
        }

        public async Task<int> CrearOrdenServicio_Cabecera(string usuario, int transportista)
        {
            int id = 0;
            string query = "DECLARE @secuencia INT = ISNULL((SELECT MAX(Secuencia) + 1 FROM TLOG_PLAN_ORDEN_SERVICIO_H " +
                "WHERE CONVERT(VARCHAR(8), Fecha, 112)= CONVERT(VARCHAR(8), GETDATE(), 112)), 1) " +
                "DECLARE @CANTIDAD INT = CONVERT(VARCHAR(8), SYSDATETIME(), 112) " +
                "INSERT INTO TLOG_PLAN_ORDEN_SERVICIO_H (FECHA, SECUENCIA, NUMERO_OS, ESTADO, USUARIO, FECHA_REGISTRO, TRANSPORTISTA) " +
                "VALUES(SYSDATETIME(), @SECUENCIA, CONCAT('OS-',@CANTIDAD,'-',@SECUENCIA),'A',@USUARIO,SYSDATETIME(),@transportista) " +
                "SELECT @@IDENTITY";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                id = await context.QueryFirstOrDefaultAsync<int>(query, new { usuario, transportista });
            }

            return id;
        }

        public async Task<List<DatosExportarSalidasDTO>> DatosExportarSalidas(DateTime? inicio, DateTime? fin)
        {
            IEnumerable<DatosExportarSalidasDTO> datos = new List<DatosExportarSalidasDTO>();
            string query = "SELECT (D.SERIE + '-' + D.NUMERO_DOCUMENTO) Guia, ISNULL(CONVERT(CHAR, d.FECHA_DOCUMENTO, 103), '') FechaGuia, D.CLIENTE_DESCRIPCION Cliente, " +
                "D.DIRECCION_DESTINO Direccion, d.DEPARTAMENTO_CLIENTE Departamento, d.FACTURA_NUMERO Comercial, d.PESO, d.BULTOS," +
                "t.DESCRIPCION Transportista, ISNULL(CONVERT(CHAR, FECHA_RETORNO, 103), '') FechaRetorno, h.NUMERO_OS OrdServicio, " +
                "ISNULL(CONVERT(CHAR, h.FECHA, 103), '') FechaServicio FROM TLOG_PLAN_ORDEN_SERVICIO_H H " +
                "INNER JOIN TLOG_PLAN_ORDEN_SERVICIO_D D ON H.ID = D.ID_H " +
                "LEFT JOIN TLO_TRANSPORTISTA T ON H.TRANSPORTISTA = T.ID " +
                "WHERE h.ESTADO = 'A' and d.ESTADO = 'A' AND CAST(h.FECHA AS Date) BETWEEN @inicio AND @fin";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                datos = await context.QueryAsync<DatosExportarSalidasDTO>(query, new { inicio, fin });
            }

            return datos.ToList();
        }

        public async Task<DatosReporteOrdenServicioPDF_DTO> DatosExportarOrdenServicio(int id)
        {
            DatosReporteOrdenServicioPDF_DTO datos = new DatosReporteOrdenServicioPDF_DTO();

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                using (SqlMapper.GridReader multi = await context.QueryMultipleAsync("usp_OrdenServicio_DatosExportarPDF", new { id }, commandType: CommandType.StoredProcedure))
                {
                    datos = multi.Read<DatosReporteOrdenServicioPDF_DTO>().FirstOrDefault();
                    datos.Detalle = multi.Read<DatosReporteOrdenServicioDetallePDF_DTO>().ToList();
                }
            
            }

            return datos;
        }

        public async Task<DatosOServicioMarcadoDTO> OrdenServicioRetornada(string ordenServicio)
        {
            DatosOServicioMarcadoDTO datos = new DatosOServicioMarcadoDTO();

            string query = "UPDATE TLOG_PLAN_ORDEN_SERVICIO_H SET FechaSalida = GETDATE() WHERE NUMERO_OS = @ordenServicio " +
                "SELECT a.NUMERO_OS OrdenServicio, ISNULL(b.DESCRIPCION, '') Transportista, " +
                "a.FECHA_REGISTRO FechaRegistro FROM TLOG_PLAN_ORDEN_SERVICIO_H a LEFT JOIN TLO_TRANSPORTISTA b " +
                "ON a.Transportista = b.ID WHERE a.NUMERO_OS = @ordenServicio";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                datos = await context.QueryFirstOrDefaultAsync<DatosOServicioMarcadoDTO>(query, new { ordenServicio });
            }

            return datos;
        }

        public async Task<(string ordenServicio, DateTime? fechaSalida)> ObtenerFechaSalidaOS(string ordenServicio)
        {
            (string ordenServicio, DateTime? fechaSalida) datos;
            string query = "SELECT NUMERO_OS OrdenServicio, FechaSalida FROM TLOG_PLAN_ORDEN_SERVICIO_H " +
                "WHERE NUMERO_OS = @ordenServicio";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                datos = await context.QueryFirstOrDefaultAsync<(string ordenServicio, DateTime? FechaSalida)>(query, new { ordenServicio });
            }

            return datos;
        }

        public async Task EliminarOrdenServicio(string ordenServicio)
        {
            string query = "DECLARE @idOs INT " +
                "SELECT @idOs = Id FROM TLOG_PLAN_ORDEN_SERVICIO_H WHERE NUMERO_OS = @ordenServicio " +
                "UPDATE TLOG_PLAN_ORDEN_SERVICIO_H SET ESTADO = 'I' WHERE Id = @idOs " +
                "UPDATE TLOG_PLAN_ORDEN_SERVICIO_D SET ESTADO = 'I' WHERE ID_H = @idOs ";

            using (SqlConnection context = new SqlConnection(_appConfig.ContextUReporteador))
            {
                await context.ExecuteAsync(query, new { ordenServicio });
            }
        }

        public async Task<List<DatosReporteGuiaOrdenServicioDTO>> DatosRptGuiasOrdenServicio(DateTime? fechaInicio, DateTime? fechaFin)
        {
            IEnumerable<DatosReporteGuiaOrdenServicioDTO> datos = new List<DatosReporteGuiaOrdenServicioDTO>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                datos = await context.QueryAsync<DatosReporteGuiaOrdenServicioDTO>("usp_satelite_RptGuiaOrdenServicio", new { fechaInicio, fechaFin }, commandType: CommandType.StoredProcedure);
            }

            return datos.ToList();
        }

    }
}
