using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using SatelliteCore.Api.ReportServices.Contracts.Actaverifacioncc;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.ReportServices.Contracts.Comercial;
using System.Linq;

namespace SatelliteCore.Api.Services
{
    public class ComercialServices : IComercialServices
    {
        private readonly IComercialRepository _comercialRepository;

        public ComercialServices(IComercialRepository comercialRepository)
        {
            _comercialRepository = comercialRepository;
        }
        public async Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos)
        {
            return await _comercialRepository.ListarCotizaciones(datos);
        }

        public async Task<FormatoCotizacionEntity> ObtenerEstructuraFormato(DatosEstructuraFormatoCotizacion datos)
        {
            return await _comercialRepository.ObtenerEstructuraFormato(datos);
        }

        public async Task<int> RegistrarRespuestas(FormatoCotizacionRespuesta datos)
        {
            return await _comercialRepository.RegistrarRespuestas(datos);
        }

        public async Task<(List<DetalleProtocoloAnalisis>, int)> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos)
        {
            return await _comercialRepository.ListarProtocoloAnalisis(datos);
        }
        public async Task<ResponseModel<string>> ListarProtocoloAnalisisExportar(DatosProtocoloAnalisisListado datos)
        {
            (List<DetalleProtocoloAnalisis> lista, int totalRegistros) result =  await _comercialRepository.ListarProtocoloAnalisis(datos);

            ReporteExcelProtocoloAnalisis ExporteProtocoloAnalisis = new ReporteExcelProtocoloAnalisis();
            string reporte = ExporteProtocoloAnalisis.GenerarReporteProtocoloAnalisis(result.lista);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);


            return Respuesta;
        }


        public async Task<List<DetalleClientes>> ListarClientes()
        {
            return await _comercialRepository.ListarClientes();
        }

        public async Task<IEnumerable<FormatoLicitaciones>> ListarDocumentoLicitacion(DatosFormatoDocumentoLicitacion dato)
        {
            return await _comercialRepository.ListarDocumentoLicitacion(dato);
        }

        public async Task<ResponseModel<string>> NumerodeGuiaLicitacion(ListarOpcionesImprimir dato)
        {
            StringBuilder builder = new StringBuilder();
            foreach (var safePrime in dato.ListaGuias)
            {
                builder.Append("'" + safePrime.GuiasNumero + "'").Append(",");
            }
            string result = builder.ToString();
            string final = result.Remove(result.Length - 1);

            FormatoReporteGuiaRemisionesModel respuesta = await _comercialRepository.NumerodeGuiaLicitacion(final);
            IEnumerable<FormatoReporteProtocoloModel> ListarProtocolo = await _comercialRepository.NumerodeGuiaProtocolo(final);

            List<DReportGuiaRemisionModel> aux = null;

            foreach (CReporteGuiaRemisionModel guia in respuesta.CabeceraReporteGuiaRemision)
            {
                aux = null;
                aux = respuesta.DetalleReporteGuiaRemision.FindAll(x => x.Guia == guia.GuiaNumero);

                if (aux.Count > 0)
                    guia.DetalleGuia.AddRange(aux);
            }

            IEnumerable<FormatoReporteProtocoloModel> aux2= null;

            foreach (DReportGuiaRemisionModel detalle in respuesta.DetalleReporteGuiaRemision)
            {
                aux2 = null;
                aux2 = ListarProtocolo.Where(x => x.OrdenFabricacion == detalle.Lote);

                if (aux2.Count() > 0)
                    detalle.DetalleProtocolo.AddRange(aux2);
            }


            if (respuesta.CabeceraReporteGuiaRemision.Count == 0)
                return new ResponseModel<string>(false, "Falta Completar Datos de la cabecera", "");

            ActaVerificacioncc actaverificacion = new ActaVerificacioncc();
            string reporte = actaverificacion.GenerarReporteActaVerificacion(respuesta.CabeceraReporteGuiaRemision,dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
        }

        public async Task<DatoPedidoDocumentoModel> NumeroPedido(string pedido)
        {
            return await _comercialRepository.NumeroPedido(pedido);
        }

        public async Task<ResponseModel<string>> RegistrarRotuladosPedido(DatosEstructuraNumeroRotuloModel dato, int idUsuario)
        {
             await _comercialRepository.RegistrarRotuladosPedido(dato, idUsuario);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "");
        }


        public async Task<IEnumerable<FormatoGuiaPorFacturarModel>> ListarGuiaporFacturar(DatosEstructuraGuiaPorFacturarModel dato)
        {
            return await _comercialRepository.ListarGuiaporFacturar(dato);
        }


        public async Task RegistrarGuiaporFacturar(DatoFormatoEstructuraGuiaFacturada dato, int idUsuario)
        {
            await _comercialRepository.RegistrarGuiaporFacturar(dato, idUsuario);
        }



        public async Task<string> ListarGuiaporFacturarExportar(DatosEstructuraGuiaPorFacturarModel dato)
        {
            string reporte = "";
            IEnumerable<FormatoGuiaPorFacturarModel> listaGuiaPorFacturar = new List<FormatoGuiaPorFacturarModel>();
            listaGuiaPorFacturar = await _comercialRepository.ListarGuiaporFacturar(dato);
            ReporteGuiaporFacturar GuiaFactura = new ReporteGuiaporFacturar();
            reporte = GuiaFactura.ExportarListarGuiaPorFactura(listaGuiaPorFacturar, dato);
            return reporte;
        }



    }
}
