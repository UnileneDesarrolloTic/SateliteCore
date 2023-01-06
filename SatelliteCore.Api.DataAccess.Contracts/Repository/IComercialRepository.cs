using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Report.Comercial;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IComercialRepository
    {
        public Task<FormatoCotizacionEntity> ObtenerEstructuraFormato(DatosEstructuraFormatoCotizacion datos);
        //public Task<int> RegistrarLote(LoteEntity lote);
        public Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos);
        public Task<int> RegistrarRespuestas(FormatoCotizacionRespuesta datos);
        public Task<List<DetalleProtocoloAnalisis>> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos);
        public Task<List<DetalleClientes>> ListarClientes();
        public Task<IEnumerable<FormatoLicitaciones>> ListarDocumentoLicitacion(DatosFormatoDocumentoLicitacion dato);
        public Task<FormatoReporteGuiaRemisionesModel> NumerodeGuiaLicitacion(string dato);
        public Task<IEnumerable<FormatoReporteProtocoloModel>> NumerodeGuiaProtocolo(string dato);
        public Task<DatoPedidoDocumentoModel> NumeroPedido(string pedido);
        public Task RegistrarRotuladosPedido(DatosEstructuraNumeroRotuloModel dato, int idUsuario);
        public Task<IEnumerable<FormatoGuiaPorFacturarModel>> ListarGuiaporFacturar(DatosEstructuraGuiaPorFacturarModel dato);
        public Task<IEnumerable<FormatoGuiaPorFacturarGeneralModel>> ListarGuiaporFacturarGeneral(DatosEstructuraGuiaPorFacturarModel dato);
        public Task RegistrarGuiaporFacturar(DatoFormatoEstructuraGuiaFacturada dato, int idUsuario);
        public Task<List<DetalleProtocoloAnalisis>> ListaProtocolosPorFacturaOPedido(DatosProtocoloAnalisisListado datos);
        public Task<List<DetalleProtocoloAnalisis>> ListaProtocolosPorGuiaRemision(DatosProtocoloAnalisisListado datos);
        public Task<List<DetalleProtocoloAnalisis>> ListaProtocolosSinTipoDocumento(string ordenFabricacion, string lote);
        public Task<List<DetalleProtocoloAnalisis>> ListaProtocolosPorCotizacion(string numeroDocumento, int idCliente, DateTime? fechaInicio, DateTime? fechaFin);
        public Task<(List<ProtocoloCabeceraModel> cabeceras, List<ProtocoloDetalleModel> detalles)> ObtenerDatosReporteProtocolo(string ordenFabricacion);
        public Task<List<ValidacionProtocoloDTO>> ValidarSiTieneProtocolo_OF(string ordenesFabricacion);

    }
}
