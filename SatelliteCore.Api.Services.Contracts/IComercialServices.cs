using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IComercialServices
    {
        public Task<FormatoCotizacionEntity> ObtenerEstructuraFormato(DatosEstructuraFormatoCotizacion datos);
        //public Task<int> RegistrarLote(LoteEntity lote);
        public Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos);
        public Task<int> RegistrarRespuestas(FormatoCotizacionRespuesta datos);
        public Task<(List<DetalleProtocoloAnalisis>, int)> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos);
        public Task<ResponseModel<string>> ListarProtocoloAnalisisExportar(DatosProtocoloAnalisisListado datos);
        public Task<List<DetalleClientes>> ListarClientes();
        public Task<IEnumerable<FormatoLicitaciones>> ListarDocumentoLicitacion(DatosFormatoDocumentoLicitacion dato);
        public Task<ResponseModel<string>> NumerodeGuiaLicitacion(ListarOpcionesImprimir dato);
        public Task<DatoPedidoDocumentoModel> NumeroPedido(string pedido);
        public Task<ResponseModel<string>> RegistrarRotuladosPedido(DatosEstructuraNumeroRotuloModel dato,int idUsuario);
        public Task<IEnumerable<FormatoGuiaPorFacturarModel>> ListarGuiaporFacturar(DatosEstructuraGuiaPorFacturarModel dato);
        public Task RegistrarGuiaporFacturar(DatoFormatoEstructuraGuiaFacturada dato, int idUsuario);
        public Task<string> ListarGuiaporFacturarExportar(DatosEstructuraGuiaPorFacturarModel dato);

    }
}
