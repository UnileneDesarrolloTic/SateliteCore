using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts
{
    public interface IAnalisisMateriaPrimaRepository
    {
        public Task<IEnumerable<ListaAnalisisMateriaPrimaDTO>> ListaOrdenesCompra(string ordenCompra, string codigoAnalisis);
        public Task CrearAnalisisHebra(GuardarAnalisisHebraDTO analisis);
        public Task<bool> ValidarSiExisteAnalisisHebra(string ordenCompra, string numeroAnalisis);
        public Task ActualizarAnalisisHebra(GuardarAnalisisHebraDTO analisis);
        public Task<AnalisisHebraDatosGeneralesDTO> DatosGeneralesAnalisisHebra(string ordenCompra, string numeroAnalisis);
        public Task<List<PlantillaDetalleProtocoloDTO>> DatosProtocoloAnalisis(string ordenCompra, string numeroAnalisis);
        public Task GuardarDatosProtocoloMateriPrima(List<GuardarProtocoloMateriaPrimaDTO> protocolo);
        public Task<bool> ValidarRegistroAnalisisHebra(string ordenCompra, string numeroAnalisis);
        public Task<PlantillaProtocoloDTO> DatosReporteProtocolo(string ordenCompra, string numeroAnalisis);
    }
}
