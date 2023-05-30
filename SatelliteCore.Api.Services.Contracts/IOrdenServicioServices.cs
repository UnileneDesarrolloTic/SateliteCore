using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IOrdenServicioServices
    {
        public Task<ResponseModel<IEnumerable<ListarOrdenServicioResponseDTO>>> ListarOrdenServicio(DateTime fechaInicio, DateTime fechaFin);
        public Task<ResponseModel<IEnumerable<ListaTransportistaComboxResponse>>> ListarTransportistaCombox();
        public Task<ResponseModel<IEnumerable<DetalleOrdenServicioResponse>>> ListaDetalleOrdenServicio(int codigoOrdenServicio);
        public Task<ResponseModel<IEnumerable<OrdenServicioGuiaRemisionResponse>>> ListaGuiaRemision(DateTime fechaInicio, DateTime fechaFin);
        public Task<ResponseModel<string>> ModificarOrdenServicio (OrdenServicioModificadosDTO ordenes, string usuario);
        public Task<ResponseModel<string>> EliminarDetalleOrdenServicio(int id);
        public Task<ResponseModel<string>> EditarGuiaRemision(EditarGuiaOS_DTO datosGuia);
        public Task<ResponseModel<string>> GuardarTransportista(DatosTransportistaDTO datosTransportista);
        public Task<ResponseModel<string>> NuevaOrdenServicio(DatosRegistrarOrdenServicioDTO ordenServicio);
        public Task<ResponseModel<string>> ExportarSalidas(DateTime? inicio, DateTime? fin);
        public Task<ResponseModel<string>> ExportarOrdenServicio(int id);
        public Task<ResponseModel<DatosOServicioMarcadoDTO?>> OrdenServicioRetornada(string ordenServicio);
        public Task<ResponseModel<string>> EliminarOrdenServicio(string ordenServicio);
        public Task<ResponseModel<string>> ReporteGuiaOrdenServicio(DateTime? fechaInicio, DateTime? fechaFin);
    }
}
