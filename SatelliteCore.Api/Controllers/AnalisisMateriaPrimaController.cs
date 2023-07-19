using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class AnalisisMateriaPrimaController : ControllerBase
    {
        private readonly IAnalisisMateriaPrimaServices _analisisMateriaPrimaServices;
        private readonly IAppConfig _appConfig;
        public AnalisisMateriaPrimaController(IAnalisisMateriaPrimaServices analisisMateriaPrimaServices, IAppConfig appConfig)
        {
            _analisisMateriaPrimaServices = analisisMateriaPrimaServices;
            _appConfig = appConfig;
        }

        [HttpGet("listarAnalisisMateriaPrima")]
        public async Task<IActionResult> ListarCertificados(string ordenCompra, string codigoAnalisis)
        {
            ResponseModel<IEnumerable<ListaAnalisisMateriaPrimaDTO>> listaAnalisis = await _analisisMateriaPrimaServices.ListaAnalisisMateriaPrima(ordenCompra, codigoAnalisis);
            return Ok(listaAnalisis);
        }

        [HttpPost("guardarAnalisisHebra")]
        public async Task<IActionResult> GuardarAnalisisHebra(GuardarAnalisisHebraDTO analisis)
        {
            ResponseModel<string> result = await _analisisMateriaPrimaServices.GuardarAnalisisHebra(analisis);
            return Ok(result);
        }

        [HttpGet("datosGeneralesAnalisisHebra")]
        public async Task<IActionResult> DatosGeneralesAnalisisHebra(string ordenCompra, string numeroAnalisis)
        {
            ResponseModel<AnalisisHebraDatosGeneralesDTO> result = await _analisisMateriaPrimaServices.DatosGeneralesAnalisisHebra(ordenCompra, numeroAnalisis);
            return Ok(result);
        }

        [HttpGet("rptAnalisisMPHebra")]
        public async Task<IActionResult> RptAnalisisMateriaPrimaHebra(string ordenCompra, string numeroAnalisis)
        {
            ResponseModel<string> result = await _analisisMateriaPrimaServices.RptAnalisisMateriaPrimaHebra(ordenCompra, numeroAnalisis);
            return Ok(result);
        }

        [HttpGet("datosProtocolo")]
        public async Task<IActionResult> DatosProtocolo(string ordenCompra, string numeroAnalisis)
        {
            ResponseModel<List<PlantillaDetalleProtocoloDTO>> protocolo = await _analisisMateriaPrimaServices.DatosProtocoloAnalisis(ordenCompra, numeroAnalisis);
            return Ok(protocolo);
        }

        [HttpPost("guardarDatosProtocoloMateriaPrima")]
        public async Task<IActionResult> GuardarDatosProtocoloMateriPrima(List<GuardarProtocoloMateriaPrimaDTO> protocolo)
        {
            ResponseModel<string> result = await _analisisMateriaPrimaServices.GuardarDatosProtocoloMateriPrima(protocolo);
            return Ok(result);
        }

        [HttpGet("rptProtocoloAnalisisMateriaPrima")]
        public async Task<IActionResult> RptProtocoloAnalisisMateriaPrima(string ordenCompra, string numeroAnalisis)
        {
            ResponseModel<string> result = await _analisisMateriaPrimaServices.RptProtocoloAnalisisMateriaPrima(ordenCompra, numeroAnalisis);
            return Ok(result);
        }
    }
}
