﻿using SatelliteCore.Api.DataAccess.Contracts.Repository;
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
using SystemsIntegration.Api.Models.Exceptions;
using SatelliteCore.Api.Models.Report.Comercial;
using SatelliteCore.Api.Models.Request.Comercial;
using System;
using System.Globalization;

namespace SatelliteCore.Api.Services
{
    public class ComercialServices : IComercialServices
    {
        private readonly IComercialRepository _comercialRepository;
        private readonly IControlCalidadRepository _controlCalidadRepository;

        public ComercialServices(IComercialRepository comercialRepository, IControlCalidadRepository controlCalidadRepository)
        {
            _comercialRepository = comercialRepository;
            _controlCalidadRepository = controlCalidadRepository;
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

        public async Task<ResponseModel<List<DetalleProtocoloAnalisis>>> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos)
        {
            ValidarDatosParaListarProtocolos(datos);

            List<DetalleProtocoloAnalisis> listaProtocolos = await FiltrarProtocolosPorFiltroTipoDocumento(datos);

            if(listaProtocolos == null)
                return new ResponseModel<List<DetalleProtocoloAnalisis>>(true, Constante.MESSAGE_SUCCESS, listaProtocolos);

            string ordenesFabricacion = null;

            listaProtocolos.ForEach(x =>
            {
                if (!string.IsNullOrEmpty(x.OrdenFabricacion))
                    ordenesFabricacion = ordenesFabricacion + x.OrdenFabricacion + ",";
            });

            if (!string.IsNullOrEmpty(ordenesFabricacion))
            {
                List<ValidacionProtocoloDTO> validacion = await _comercialRepository.ValidarSiTieneProtocolo_OF(ordenesFabricacion);

                if (validacion == null)
                    return new ResponseModel<List<DetalleProtocoloAnalisis>>(true, Constante.MESSAGE_SUCCESS, listaProtocolos);

                int cantidadRegistros = listaProtocolos.Count - 1;

                for (int i = 0; i <= cantidadRegistros; i++)
                {
                    DetalleProtocoloAnalisis protocolo = listaProtocolos[i];
                    string tieneProtocolo = validacion.FirstOrDefault(x => x.OrdenFabricacion == protocolo.OrdenFabricacion)?.Tiene ?? "No";

                    protocolo.TieneProtocolo = tieneProtocolo;
                    listaProtocolos[i] = protocolo;
                }
            }

            if (!string.IsNullOrEmpty(datos.Protocolo))
                listaProtocolos = listaProtocolos.FindAll(x => x.TieneProtocolo == datos.Protocolo);

            return new ResponseModel<List<DetalleProtocoloAnalisis>>(true, Constante.MESSAGE_SUCCESS, listaProtocolos);
        }

        public async Task<ResponseModel<string>> ListarProtocoloAnalisisExportar(DatosProtocoloAnalisisListado datos)
        {
            ValidarDatosParaListarProtocolos(datos);

            List<DetalleProtocoloAnalisis> listaProtocolo = await FiltrarProtocolosPorFiltroTipoDocumento(datos);

            if (listaProtocolo.Count < 1)
                return new ResponseModel<string>(false, "No se encontro protocolo para exportar.", null);

            ReporteExcelProtocoloAnalisis ExporteProtocoloAnalisis = new ReporteExcelProtocoloAnalisis();
            string reporte = ExporteProtocoloAnalisis.GenerarReporteProtocoloAnalisis(listaProtocolo);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
        }

        public async Task<ResponseModel<string>> GenerarReporteProtocoloAnalisis(int idioma, List<string> ordenesFabricacion)
        {
            if (ordenesFabricacion == null)
                throw new ValidationModelException("Las Ordenes de fabricación no válidos.");

            if (idioma <= 0)
                throw new ValidationModelException("El idioma no válido");

            string cadenaOrdenFabricación = null;
            string mensajeReturn = Constante.MESSAGE_SUCCESS;

            ordenesFabricacion.ForEach(l =>
            {
                if (string.IsNullOrEmpty(l))
                    throw new ValidationModelException("Ordenes de fabricación no válidos.");

                cadenaOrdenFabricación = cadenaOrdenFabricación + l + ",";
            });

            (List<ProtocoloCabeceraModel> cabecerasPro, List<ProtocoloDetalleModel> detallesPro) 
                = await _comercialRepository.ObtenerDatosReporteProtocolo(idioma, cadenaOrdenFabricación);
            string versionProtocolo = await _controlCalidadRepository.VersionProtocolo();

            if (cabecerasPro.Count < 1)
            {
                mensajeReturn = "No se encontro datos de los protocolos solicitados.";
                return new ResponseModel<string>(false, mensajeReturn, null);
            }


            if (cabecerasPro.Count < ordenesFabricacion.Count)
                mensajeReturn = "No se pudo encontrar datos de algunas Ord. Fabricación enviada.";

            
            List<ProtocoloReportModel> listaProtocolo = new List<ProtocoloReportModel>();

            foreach (ProtocoloCabeceraModel cabecera in cabecerasPro)
            {
                List<ProtocoloDetalleModel> detalle = detallesPro.FindAll(x => x.OrdenFabricacion == cabecera.OrdenFabricacion);

                if(detalle.Count < 1)
                    throw new ValidationModelException($"No se pudo encontrar el detalle del OF: {cabecera.OrdenFabricacion}");

                ProtocoloReportModel datoProtocolo = new ProtocoloReportModel();
                datoProtocolo.Cabecera = cabecera;
                datoProtocolo.Detalle = detallesPro.FindAll(x => x.OrdenFabricacion == cabecera.OrdenFabricacion);

                listaProtocolo.Add(datoProtocolo);
            }

            ReportePdfProtocoloAnalisis protocolo = new ReportePdfProtocoloAnalisis(idioma, listaProtocolo, versionProtocolo);
            string reporte = protocolo.GenerarReporte();

            return new ResponseModel<string>(true, mensajeReturn, reporte);
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
            string versionProtocolo = await _controlCalidadRepository.VersionProtocolo();

            if (respuesta.CabeceraReporteGuiaRemision.Count == 0)
                return new ResponseModel<string>(false, "Falta Completar Datos de la cabecera", "");

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




            ActaVerificacioncc actaverificacion = new ActaVerificacioncc();
            string reporte = actaverificacion.GenerarReporteActaVerificacion(respuesta.CabeceraReporteGuiaRemision,dato, versionProtocolo);

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

        public async Task<ResponseModel<RegistroRecepcionGuiaResponseDTO?>> RegistrarAdministracionGuia(string numeroDocumento, int idUsuario, string tipoRegistro)
        {
            if(string.IsNullOrWhiteSpace(numeroDocumento) || string.IsNullOrWhiteSpace(tipoRegistro))
                throw new ValidationModelException();

            if(numeroDocumento.IndexOf("-") == -1 || !numeroDocumento.StartsWith("T", true, new CultureInfo("en-US")) || idUsuario < 1 || (tipoRegistro != "I" && tipoRegistro != "O"))
                throw new ValidationModelException();

            int separador = numeroDocumento.IndexOf("-");
            string serieNumero = numeroDocumento.Substring(0, separador);
            string documentoNumero = numeroDocumento.Substring(separador + 1).PadLeft(10, '0');

            (string guiaNumero, string comentariosEntrega, DateTime? FechaEnvioAlmacen) datosValidacion = await _comercialRepository.ObtenerGuiaRegistrada(serieNumero, documentoNumero);

            if (string.IsNullOrWhiteSpace(datosValidacion.guiaNumero))
                return new ResponseModel<RegistroRecepcionGuiaResponseDTO?>(false, $"No se pudo encontrar la guía ({serieNumero.ToUpper() + "-" + documentoNumero}).", null);

            if (tipoRegistro == "I")
            {
                if (datosValidacion.FechaEnvioAlmacen is null)
                    return new ResponseModel<RegistroRecepcionGuiaResponseDTO?>(false, $"La guía ({serieNumero.ToUpper() + "-" + documentoNumero}) no tiene fecha de envío a almacen.", null);

                if (datosValidacion.comentariosEntrega == "1")
                    return new ResponseModel<RegistroRecepcionGuiaResponseDTO?>(false, $"La guía ({serieNumero.ToUpper() + "-" + documentoNumero}) ya se encuentra registrada como recepcionada.", null);
            }
            else
            {
                if (datosValidacion.FechaEnvioAlmacen != null)
                    return new ResponseModel<RegistroRecepcionGuiaResponseDTO?>(false, $"Ya se registro el envío de la guía ({serieNumero.ToUpper() + "-" + documentoNumero}).", null);
            }

            RegistroRecepcionGuiaResponseDTO datosGuia = await _comercialRepository.RegistrarAdministracionGuia(serieNumero, documentoNumero, idUsuario, tipoRegistro);

            return new ResponseModel<RegistroRecepcionGuiaResponseDTO?>(datosGuia);
        }

        public async Task<string> ListarGuiaporFacturarExportar(DatosEstructuraGuiaPorFacturarModel dato)
        {
            IEnumerable<FormatoGuiaPorFacturarModel> listaGuiaPorFacturar = await _comercialRepository.ListarGuiaporFacturar(dato);

            ReporteGuiaporFacturar GuiaFactura = new ReporteGuiaporFacturar();
            string reporte = GuiaFactura.ExportarListarGuiaPorFactura(listaGuiaPorFacturar, dato);
            return reporte;

        }

        private void ValidarDatosParaListarProtocolos(DatosProtocoloAnalisisListado datos)
        {
            if (string.IsNullOrEmpty(datos.Lote) && string.IsNullOrEmpty(datos.OrdenFabricacion) && string.IsNullOrEmpty(datos.TipoDocumento))
                throw new ValidationModelException("Debe de contar como mínimo con los filtros: Ord. Fabricacion, Lote ó Documento");

            if (!string.IsNullOrEmpty(datos.TipoDocumento) && string.IsNullOrEmpty(datos.NumeroDocumento))
                throw new ValidationModelException("Debe de ingresar el tipo y número de documento");

            if ((datos.FechaInicio == null && datos.FechaFin != null) || (datos.FechaInicio != null && datos.FechaFin == null))
                throw new ValidationModelException("Las fechas no son válidas.");

            if (datos.FechaInicio != null && datos.FechaFin != null)
            {
                if (datos.FechaInicio > datos.FechaFin)
                    throw new ValidationModelException("La fecha inicio no puede ser mayor a la fecha fin.");

                if (datos.FechaInicio?.ToString("MM/yyyy") != datos.FechaFin?.ToString("MM/yyyy"))
                    throw new ValidationModelException("Las fechas deben pertenecer al mismo periodo");
            }
        }

        private async Task<List<DetalleProtocoloAnalisis>> FiltrarProtocolosPorFiltroTipoDocumento(DatosProtocoloAnalisisListado datos)
        {
            List<DetalleProtocoloAnalisis> listaProtocolos = new List<DetalleProtocoloAnalisis>();

            if (datos.TipoDocumento == "")
                listaProtocolos = await _comercialRepository.ListaProtocolosSinTipoDocumento(datos.OrdenFabricacion, datos.Lote);

            if (datos.TipoDocumento == "P" || datos.TipoDocumento == "F")
            {
                datos.NumeroDocumento = $"%{datos.NumeroDocumento.Trim()}";
                listaProtocolos = await _comercialRepository.ListaProtocolosPorFacturaOPedido(datos);
            }

            if (datos.TipoDocumento == "G")
            {
                datos.NumeroDocumento = $"%{datos.NumeroDocumento.Trim()}";
                listaProtocolos = await _comercialRepository.ListaProtocolosPorGuiaRemision(datos);
            }


            if (datos.TipoDocumento == "C")
            {
                datos.NumeroDocumento = datos.NumeroDocumento.Trim().PadLeft(10, '0');
                listaProtocolos = await _comercialRepository.ListaProtocolosPorCotizacion(datos.NumeroDocumento, datos.IdCliente, datos.FechaInicio, datos.FechaFin);
            }

            return listaProtocolos;
        }

    }
}
