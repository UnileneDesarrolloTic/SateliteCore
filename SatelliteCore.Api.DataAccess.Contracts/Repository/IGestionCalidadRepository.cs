using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request.GestionCalidad;
using SatelliteCore.Api.Models.Request.GestorDocumentario;
using SatelliteCore.Api.Models.Response.GestionCalidad;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IGestionCalidadRepository
    {
        public Task<List<MateriaPrimaDTO>> ObtenerMateriaPrima(string tipoConsulta, string lote);
        public Task<List<OrdenCompraPorLoteDTO>> OrdenCompraPorlote(List<string> lotes);
        public Task<List<OrdenFabricacionPorLoteDTO>> OrdenFabricacionPorlotes(List<string> lotes);
        public Task<List<DocumentosPedidosDTO>> OrdenDocumentosPedidosPorLotes(List<string> lotes);
        public Task<List<GuiaDocumentos>> OrdenGuiasRelacionadasPorLotes(List<string> lotes);
        public Task<List<VentasPorClienteDTO>> VentasPorCliente(RequestFiltroVentaCliente filtros);
        public Task<(List<ListaReclamosDTO>, int cantidadRegistros)> ListarReclamosQuejas(FiltrosListaReclamosDTO filtros);
        public Task RegistrarReclamo(TBMReclamosEntity reclamo);
        public Task<CabeceraDetalleReclamoDTO> ObtenerCabeceraDetalleReclamo(string codigoReclamo);
        public Task<IEnumerable<LotesFiltradosReclamo>> LotesFiltradosReclamo(FiltrosLotesReclamosDTO filtros);
        public Task<DatosLoteReclamoDTO> DatosItemLote(string lote);
        public Task<int> GuardarDetalleReclamo(TBDReclamosEntity detalle);
        public Task<bool> ValidarExisteDetalleReclamo(string codReclamo, string lote, string ordenFabricacion, string tipoDocumento, string documento);
        public Task<List<DetalleReclamoDTO>> ListarDetalleReclamo(string codReclamo);
        public Task<CabeceraReclamoLoteDTO> CabeceraReclamoLote(string codReclamo, string lote, string documento);
        public Task<int> ActualizarDetalleLoteReclamo(TBDReclamosEntity detalle);
        public Task<int> RegistrarRespuestaReclamo(RespuestaReclamoDTO respuesta);
        public Task Validar_ActualizarEstadoReclamo(int idDetalle);
        public Task<string> ObtenerEstadoLoteDetalleReclamo(int idDetalle);
        public Task<IEnumerable<DatosFormatoListarSsomaModel>> ListarSsoma(int TipoDocumento, string Codigo, int Estado);
        public Task<object> RegistrarSsoma(DatosFormatoRegistrarSsomaModel dato,string UsuarioSesion);
        public Task<int> EliminarSsoma(int idSsoma, string UsuarioSesion);
    }
}
