
using SatelliteCore.Api.DataAccess.Contracts;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using SatelliteCore.Api.ReportServices.Contracts.AnalisisMateriaPrima.General;
using SatelliteCore.Api.ReportServices.Contracts.AnalisisMateriaPrima.Hebra;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class AnalisisMateriaPrimaServices : IAnalisisMateriaPrimaServices
    {
        private readonly IAnalisisMateriaPrimaRepository _analisisMateriaPrimaRepository;
        public AnalisisMateriaPrimaServices(IAnalisisMateriaPrimaRepository analisisMateriaPrimaRepository)
        {
            _analisisMateriaPrimaRepository = analisisMateriaPrimaRepository;
        }

        public async Task<ResponseModel<IEnumerable<ListaAnalisisMateriaPrimaDTO>>> ListaAnalisisMateriaPrima(string ordenCompra, string codigoAnalisis)
        {
            if (string.IsNullOrWhiteSpace(ordenCompra) && string.IsNullOrWhiteSpace(codigoAnalisis))
                throw new ValidationModelException();

            if (!string.IsNullOrEmpty(ordenCompra))
            {
                ordenCompra = "000000" + ordenCompra;

                if (ordenCompra.Substring(0, 3) != "FOR")
                    ordenCompra = "FOR" + ordenCompra.Substring((ordenCompra.Length - 6), 6);
            }

            IEnumerable<ListaAnalisisMateriaPrimaDTO> listaAnaliss = await _analisisMateriaPrimaRepository.ListaOrdenesCompra(ordenCompra, codigoAnalisis);

            return new ResponseModel<IEnumerable<ListaAnalisisMateriaPrimaDTO>>(listaAnaliss);
        }

        public async Task<ResponseModel<string>> GuardarAnalisisHebra(GuardarAnalisisHebraDTO analisis)
        {
            if (!analisis.Cabecera.ValidarDatos())
                throw new ValidationModelException("Los datos de la cebecera no son válidos.");

            if (analisis.Detalle.Count != 20)
                throw new ValidationModelException("Los datos del detalle no son válidos");

            analisis.Detalle.ForEach(x =>
            {
                if (!x.ValidarDatos())
                    throw new ValidationModelException("Los datos del detalle no son válidos.");
            });

            bool existeAnalisis = await _analisisMateriaPrimaRepository.ValidarSiExisteAnalisisHebra(analisis.Cabecera.OrdenCompra, analisis.Cabecera.NumeroAnalisis);

            if (existeAnalisis)
                await _analisisMateriaPrimaRepository.ActualizarAnalisisHebra(analisis);
            else
                await _analisisMateriaPrimaRepository.CrearAnalisisHebra(analisis);

            return new ResponseModel<string>("Se ha guardado los datos del análisis.");

        }

        public async Task<ResponseModel<AnalisisHebraDatosGeneralesDTO>> DatosGeneralesAnalisisHebra(string ordenCompra, string numeroAnalisis)
        {
            if (string.IsNullOrWhiteSpace(ordenCompra) || string.IsNullOrWhiteSpace(numeroAnalisis))
                throw new ValidationModelException();

            AnalisisHebraDatosGeneralesDTO datosAnalisis = await _analisisMateriaPrimaRepository.DatosGeneralesAnalisisHebra(ordenCompra, numeroAnalisis);

            return new ResponseModel<AnalisisHebraDatosGeneralesDTO>(datosAnalisis);
        }

        public async Task<ResponseModel<string>> RptAnalisisMateriaPrimaHebra(string ordenCompra, string numeroAnalisis)
        {
            if (string.IsNullOrWhiteSpace(ordenCompra) || string.IsNullOrWhiteSpace(numeroAnalisis))
                throw new ValidationModelException();

            ResponseModel<AnalisisHebraDatosGeneralesDTO> datosReporte = await DatosGeneralesAnalisisHebra(ordenCompra, numeroAnalisis);
            
            if (datosReporte.Content.Detalle.Count < 1)            
                return new ResponseModel<string>(false, "No se encontrol datos del análisis", null);

            RptAnalisisMateriaPrima_PDF reporteAnalisis = new RptAnalisisMateriaPrima_PDF();
            string reporte = reporteAnalisis.GenerarReporte(datosReporte.Content);

            return new ResponseModel<string>(reporte);
        }

        public async Task<ResponseModel<List<PlantillaDetalleProtocoloDTO>>> DatosProtocoloAnalisis(string ordenCompra, string numeroAnalisis)
        {
            if (string.IsNullOrWhiteSpace(ordenCompra) || string.IsNullOrWhiteSpace(numeroAnalisis))
                throw new ValidationModelException();

            bool existeAnalisisHebra = await _analisisMateriaPrimaRepository.ValidarRegistroAnalisisHebra(ordenCompra, numeroAnalisis);
            List<PlantillaDetalleProtocoloDTO> protocolo = new List<PlantillaDetalleProtocoloDTO>();

            if (!existeAnalisisHebra)
                return new ResponseModel<List<PlantillaDetalleProtocoloDTO>>(false, "Debe de registrar el análisis de la hebra.", protocolo);

            protocolo = await _analisisMateriaPrimaRepository.DatosProtocoloAnalisis(ordenCompra, numeroAnalisis);

            return new ResponseModel<List<PlantillaDetalleProtocoloDTO>>(protocolo);

        }

        public async Task<ResponseModel<string>> GuardarDatosProtocoloMateriPrima(List<GuardarProtocoloMateriaPrimaDTO> protocolo)
        {
            if (protocolo.Count < 1)
                throw new ValidationModelException();

            await _analisisMateriaPrimaRepository.GuardarDatosProtocoloMateriPrima(protocolo);

            return new ResponseModel<string>("Se ha guardado los datos del protocolo");
        }

        public async Task<ResponseModel<string>> RptProtocoloAnalisisMateriaPrima(string ordenCompra, string numeroAnalisis)
        {
            if (string.IsNullOrWhiteSpace(ordenCompra) || string.IsNullOrWhiteSpace(numeroAnalisis))
                throw new ValidationModelException();

            PlantillaProtocoloDTO datosReporte = await _analisisMateriaPrimaRepository.DatosReporteProtocolo(ordenCompra, numeroAnalisis);

            if (string.IsNullOrWhiteSpace(datosReporte.Cabecera.Analisis) ||  datosReporte.Detalle.Count < 1)
                return new ResponseModel<string>(false, "No se encontro datos del protocolo", null);

            ProtocoloMateriaPrima_PDF reporteProtocolo = new ProtocoloMateriaPrima_PDF();
            string reporte = reporteProtocolo.GenerarReporte(datosReporte);

            return new ResponseModel<string>(reporte);
        }




    }
}
