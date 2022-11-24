using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
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
    }
}
