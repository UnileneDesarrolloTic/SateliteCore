﻿using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.AnalsisAguja;
using SatelliteCore.Api.ReportServices.Contracts.Logistica;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class LogisticaServices : ILogisticaServices
    {
        private readonly ILogisticaRepository _logisticaRepository;

        public LogisticaServices(ILogisticaRepository logisticaRepository)
        {
            _logisticaRepository = logisticaRepository;
        }
        public async Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string NumeroGuia)
        {
            return await _logisticaRepository.ObtenerNumeroGuias(NumeroGuia);
        }

        public async Task<ResponseModel<string>> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato)
        {
             await _logisticaRepository.RegistrarRetornoGuia(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos");
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


    }
}
