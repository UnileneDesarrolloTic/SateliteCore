using SatelliteCore.Api.Models.Dto.GestionCalidad;
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
        public Task<IEnumerable<DatosFormatoListarSsomaModel>> ListarSsoma(int TipoDocumento, string Codigo, int Estado);
        public Task<object> RegistrarSsoma(DatosFormatoRegistrarSsomaModel dato,string UsuarioSesion);
        public Task<int> EliminarSsoma(int idSsoma, string UsuarioSesion);
    }
}
