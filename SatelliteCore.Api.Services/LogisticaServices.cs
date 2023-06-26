using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using SatelliteCore.Api.ReportServices.Contracts.Logistica;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class LogisticaServices : ILogisticaServices
    {
        private readonly ILogisticaRepository _logisticaRepository;
        private readonly ICommonRepository _commonRepository;

        public LogisticaServices(ILogisticaRepository logisticaRepository, ICommonRepository commonRepository)
        {
            _logisticaRepository = logisticaRepository;
            _commonRepository = commonRepository;
        }
        public async Task<ResponseModel<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string numeroguia, string Usuario)
        {
            if (string.IsNullOrEmpty(numeroguia))
                throw new ValidationModelException("Debe ingresar información completa");

            DatosFormatoPlanOrdenServicosD respuesta = new DatosFormatoPlanOrdenServicosD();
            respuesta = await _logisticaRepository.ObtenerNumeroGuias(numeroguia);

            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.LO_GESTION_GUIA_BUSCAR;
            evento.Usuario = Usuario;
            evento.Opcional = numeroguia;
            await _commonRepository.RegistroLogEvento(evento);


            if (respuesta.NumeroGuia == null)
                return new ResponseModel<DatosFormatoPlanOrdenServicosD>(false, "No hay información de esa guia" , respuesta);
            return new ResponseModel<DatosFormatoPlanOrdenServicosD>(true, Constante.MESSAGE_SUCCESS, respuesta);
        }

        public async Task<ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>>> ListarRetornoGuia(DatosFormatoGestionGuiasClienteModel dato)
        {
            IEnumerable<DatosFormatosReporteRetornoGuia> result = new List<DatosFormatosReporteRetornoGuia>();

            /*if (string.IsNullOrEmpty(dato.cliente) && string.IsNullOrEmpty(dato.destino) && dato.transportista == 0 && dato.activarFecha == false)
                return new ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>>(false, "Debe ingresar un valor para lograr la busqueda", result);*/

            if (dato.activarFecha)
                if (string.IsNullOrEmpty(dato.fechaInicio.ToString()) || string.IsNullOrEmpty(dato.fechaFin.ToString()))
                    return new ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>>(false, "Debe ingresar las dos fechas: inicio y fin", result);
            
            result = await _logisticaRepository.ListarRetornoGuia(dato);

            if(result.Count() == 0)
               return new ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>>(false,"No hay datos para mostrar", result);

            return new ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>> (true, Constante.MESSAGE_SUCCESS, result);
        }


        public async Task<ResponseModel<string>> ExportarExcelRetornoGuia(DatosFormatoGestionGuiasClienteModel dato)
        {
            IEnumerable<DatosFormatosReporteRetornoGuia> result = new List<DatosFormatosReporteRetornoGuia>();
            result = await _logisticaRepository.ListarRetornoGuia(dato);

            if (result.Count() == 0)
                return new ResponseModel<string>(false, "No hay datos para mostrar", "");

            string reporte = "";
            ReporteRetornoGuias_Excel guiareporte = new ReporteRetornoGuias_Excel();
            reporte = guiareporte.GenerarReporte(result, dato);
           
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
        }


        public async Task<ResponseModel<string>> RegistrarRetornoGuia(DatosFormatoRetornoGuiaRequest dato, string Usuario)
        {
            await _logisticaRepository.RegistrarRetornoGuia(dato);

            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.LO_GESTION_GUIA_REGISTRAR;
            evento.Usuario = Usuario;
            evento.Opcional = dato.numeroGuia;
            await _commonRepository.RegistroLogEvento(evento);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo con éxito");
        }

        public async Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato)
        {
            return await _logisticaRepository.ListarItemVentas(dato);
        }

        public async Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item)
        {
            return await _logisticaRepository.BuscarItemVentas(Item);
        }

        public async Task<ResponseModel<string>> ListarItemVentasExportar(FormatoDatosBusquedaItemsVentas dato)
        {
            IEnumerable<DatosFormatoItemVentas> result = new List<DatosFormatoItemVentas>();
            result = await _logisticaRepository.ListarItemVentas(dato);

            ReporteItemVentas ExporteItemventas = new ReporteItemVentas();
            string reporte = ExporteItemventas.GenerarReporte(result);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle()
        {
            return await _logisticaRepository.ListarItemVentasDetalle();
        }

        public async Task<ResponseModel<string>> ListarItemVentasDetalleExportar()
        {
            IEnumerable<DatosFormatoDetalledelItemVentas> result = new List<DatosFormatoDetalledelItemVentas>();
            result = await _logisticaRepository.ListarItemVentasDetalle();

            ReportItemVentasDetalle ExporteItemventasDetalle = new ReportItemVentasDetalle();
            string reporte = ExporteItemventasDetalle.GenerarReporteDetalle(result);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato)
        {
            return await _logisticaRepository.DetalleComprometidoItem(dato);
        }

        public async Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo)
        {
            return await _logisticaRepository.BuscarNumeroPedido(NumeroDocumento, Tipo);
        }

        public async Task<IEnumerable<DatosFormatoDetalleRecetaMPLogistica>> BuscardDetalleRecetaMP(string Item, string Cantidad)
        {
            return await _logisticaRepository.BuscardDetalleRecetaMP(Item, Cantidad);
        }

    }
}
