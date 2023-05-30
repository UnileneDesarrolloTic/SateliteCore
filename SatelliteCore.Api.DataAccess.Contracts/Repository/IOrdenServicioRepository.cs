using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IOrdenServicioRepository
    {
        public Task<IEnumerable<ListarOrdenServicioResponseDTO>> ListarOrdenServicio(DateTime fechaInicio, DateTime fechaFin);
        public Task<IEnumerable<ListaTransportistaComboxResponse>> ListarTransportistaCombox();
        public Task<IEnumerable<DetalleOrdenServicioResponse>> ListaDetalleOrdenServicio(int codigoOrdenServicio);
        public Task<IEnumerable<OrdenServicioGuiaRemisionResponse>> ListaGuiaRemision(DateTime fechaInicio, DateTime fechaFin);
        public Task Modificar_Peso_Bultos(List<OrdenServicioDetalle> ordenes);
        public Task RegistrarGuias_OrdenServicio(List<RegistrarGuia_OrdenServicioDTO> datosGuia);
        public Task RegistrarObjetosExtrasEnvio_OS(List<OrdenServicioDetalle> extras);
        public Task EliminarDetalleOrdenServicio(int id);
        public Task EditarGuiaRemision(EditarGuiaOS_DTO datosGuia);
        public Task ModificarTransportista(int id, int idTransportista);
        public Task GuardarTransportista(DatosTransportistaDTO datosTransportista);
        public Task<int> CrearOrdenServicio_Cabecera(string usuario, int transportista);
        public Task<List<DatosExportarSalidasDTO>> DatosExportarSalidas(DateTime? inicio, DateTime? fin);
        public Task<DatosReporteOrdenServicioPDF_DTO> DatosExportarOrdenServicio(int id);
        public Task<DatosOServicioMarcadoDTO> OrdenServicioRetornada(string ordenServicio);
        public Task<(string ordenServicio, DateTime? fechaSalida)> ObtenerFechaSalidaOS(string ordenServicio);
        public Task EliminarOrdenServicio(string ordenServicio);
        public Task<List<DatosReporteGuiaOrdenServicioDTO>> DatosRptGuiasOrdenServicio(DateTime? fechaInicio, DateTime? fechaFin);
    }
}
