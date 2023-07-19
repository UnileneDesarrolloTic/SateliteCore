using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IAnalisisMateriaPrimaServices
    {
        public Task<ResponseModel<IEnumerable<ListaAnalisisMateriaPrimaDTO>>> ListaAnalisisMateriaPrima(string ordenCompra, string codigoAnalisis);
        public Task<ResponseModel<string>> GuardarAnalisisHebra(GuardarAnalisisHebraDTO analisis);
        public Task<ResponseModel<AnalisisHebraDatosGeneralesDTO>> DatosGeneralesAnalisisHebra(string ordenCompra, string numeroAnalisis);
        public Task<ResponseModel<string>> RptAnalisisMateriaPrimaHebra(string ordenCompra, string numeroAnalisis);
        public Task<ResponseModel<List<PlantillaDetalleProtocoloDTO>>> DatosProtocoloAnalisis(string ordenCompra, string numeroAnalisis);
        public Task<ResponseModel<string>> GuardarDatosProtocoloMateriPrima(List<GuardarProtocoloMateriaPrimaDTO> protocolo);
        public Task<ResponseModel<string>> RptProtocoloAnalisisMateriaPrima(string ordenCompra, string numeroAnalisis);
    }
}
