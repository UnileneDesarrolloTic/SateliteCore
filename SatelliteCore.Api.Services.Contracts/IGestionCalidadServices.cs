using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request.GestionCalidad;
using SatelliteCore.Api.Models.Request.GestorDocumentario;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.GestionCalidad;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IGestionCalidadServices
    {
        public Task<List<MateriaPrimaDTO>> ObtenerMateriaPrima(string tipoConsulta, string lote);
        public Task<DetalleSeguimientoLoteDTO> DetalleSeguimientoPorLote(RequestLotesDetalleDTO filtros);
        public Task<List<VentasPorClienteDTO>> VentasPorCliente(RequestFiltroVentaCliente filtros);
        public Task<ResponseModel<string>> ReporteVentasPorCliente(RequestFiltroVentaCliente filtros);
        public Task<IEnumerable<DatosFormatoListarSsomaModel>> ListarSsoma(int TipoDocumento, string Codigo, int estado);
        public Task<ResponseModel<string>> RegistrarSsoma(DatosFormatoRegistrarSsomaModel dato,string UsuarioSesion);
        public Task<ResponseModel<string>> EliminarSsoma(int idSsoma, string UsuarioSesion);
        public Task<PaginacionModel<ListaReclamosDTO>> ListarReclamosQuejas(FiltrosListaReclamosDTO filtros);
        public Task<ResponseModel<object>> RegistrarReclamoCabecera(int codigoCliente, string codigoUsuarioSesion);
        public Task<ResponseModel<ReclamoDTO>> ObtenerDetalleReclamo(string codigoReclamo);
        public Task<ResponseModel<IEnumerable<LotesFiltradosReclamo>>> LotesFiltradosReclamo(FiltrosLotesReclamosDTO filtros);
        public Task<ResponseModel<DatosLoteReclamoDTO>> DatosItemLote(string lote);
        public Task<ResponseModel<object>> GuardarDetalleReclamo(TBDReclamosEntity detalle);
        public Task<ResponseModel<CabeceraReclamoLoteDTO>> DataReclamoLote(string codReclamo, string lote, string documento);
        public Task<ResponseModel<string>> ActualizarDetalleLoteReclamo(TBDReclamosEntity detalle);
        public Task<ResponseModel<string>> RegistrarRespuestaReclamo(RespuestaReclamoDTO respuesta);
    }
}
