using SatelliteCore.Api.Models.Dto.GestionCalidad;
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
    }
}
