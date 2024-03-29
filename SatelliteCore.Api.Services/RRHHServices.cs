﻿using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Report.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.RRHH;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.RRHH;
using SatelliteCore.Api.ReportServices.Contracts.RRHH;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class RRHHServices : IRRHHServices
    {
        private readonly IRRHHRepository _rrhhRepository;
        public RRHHServices(IRRHHRepository rrhhRepository)
        {
            _rrhhRepository = rrhhRepository;
        }

        public async Task<ResponseModel<IEnumerable<ReporteAsistenciaDTO>>> ListarAsistencia(DateTime fecha, int usuarioToken)
        {
            IEnumerable<ReporteAsistenciaDTO> listaAsistencia = await _rrhhRepository.ListarAsistencia(fecha, usuarioToken);
            return new ResponseModel<IEnumerable<ReporteAsistenciaDTO>>(true, Constante.MESSAGE_SUCCESS, listaAsistencia);
        }

        public async Task<ResponseModel<string>> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera data, string usuario)
        {
            _ = await _rrhhRepository.RegistrarHorasExtras(data, usuario);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con exito");
        }

        public async Task<IEnumerable<DatosFormatoListarHorasExtrasPersona>> ListarHoraExtrasPersona(DatosFormatoFiltraHoraExtras data)
        {
            if(data.Periodo != null)
                data.Periodo = data.Periodo.Replace("-","");

            IEnumerable<DatosFormatoListarHorasExtrasPersona> response = await _rrhhRepository.ListarHoraExtrasPersona(data);
            return response;
        }

        public async Task ProcesarHorasExtrasPlanilla(string periodo)
        {
            if (string.IsNullOrEmpty(periodo) || periodo.Length  != 7)
                throw new ValidationModelException("Error en los datos enviados.");

            periodo = periodo.Replace("-", "");

            await _rrhhRepository.ProcesarHorasExtrasPlanilla(periodo);
        }

        public async Task<object> BuscarInformacionHorasExtrasPersona(int Cabecera)
        {
            (DatosFormatoHorasExtrasCabeceraModel cabecera, List <DatosFormatoHorasExtrasDetalle> detalle) = await _rrhhRepository.BuscarInformacionHorasExtrasPersona(Cabecera);
            object result = new { cabecera, detalle };

            return result;
        }

        public async Task<string> ReporteHorasExtrasGeneradas(string periodo)
        {
            if(string.IsNullOrEmpty(periodo))
                throw new ValidationModelException("Error en los datos enviados !!");

            periodo = periodo.Replace("-", "");

            List<HorasExtraExportDTO> horasExtras = await _rrhhRepository.ListarHorasExtrasGeneradas(periodo);

            HorasExtrasGeneradas_Excel classReport = new HorasExtrasGeneradas_Excel(horasExtras);
            string reporte = classReport.Exportar();

            if (string.IsNullOrEmpty(reporte))
                throw new Exception("Error al generar el reporte !!");

            return reporte;
        }

        public async Task<string> FormatoAutorizacionSobretiempo (int id)
        {

            (DatosFormatoHorasExtrasCabeceraModel cabecera, List<DatosFormatoHorasExtrasDetalle> detalle) 
                = await _rrhhRepository.BuscarInformacionHorasExtrasPersona(id);

            FormatoAutorizacionSobretiempo_PDF formatoAutoSobretiempo 
                = new FormatoAutorizacionSobretiempo_PDF(cabecera, detalle);

            string reporte = formatoAutoSobretiempo.Exportar();

            if (string.IsNullOrEmpty(reporte))
                throw new Exception("Error al generar el reporte !!");

            return reporte;
        }
        public async Task<ResponseModel<string>> AutorizacionSobretiempoPorPersona(string periodo)
        {
            if(string.IsNullOrEmpty(periodo) || periodo.Length != 7)
                throw new ValidationModelException("Error en los datos enviados !!");

            periodo = periodo.Replace("-", "");

            AutorizacionSobretiempoPersonaDTO datos = await _rrhhRepository.ListarHorasExtraExportas(periodo);

            if(datos.Cabecera.Count <= 0)
                return new ResponseModel<string>(false, "No se encontro registros para este periodo", null);
            

            AutorizacionSobretiempoPorPersona_PDF reporte = new AutorizacionSobretiempoPorPersona_PDF(datos);
            string rpt = reporte.Exportar();

            return new ResponseModel<string>(true, Constante.MESSSGE_SUCCESS_REPORT, rpt);
        }

        public async Task<IEnumerable<DatosFormatoReporteComisionVendedor>> ReporteComisionVendedor(string periodo)
        {
            if (string.IsNullOrEmpty(periodo) || periodo.Length != 8)
                throw new ValidationModelException("Error en los datos enviados !!");

            IEnumerable<DatosFormatoReporteComisionVendedor> response = await _rrhhRepository.ReporteComisionVendedor(periodo);
            return response;
        }

        public async Task<ResponseModel<string>> ReporteComisionVendedorExcel(string periodo)
        {
            if (string.IsNullOrEmpty(periodo) || periodo.Length != 8)
                throw new ValidationModelException("Error en los datos enviados !!");

            IEnumerable<DatosFormatoReporteComisionVendedor> listado = await _rrhhRepository.ReporteComisionVendedor(periodo);

            if(listado.Count() == 0)
                return new ResponseModel<string>(false, "No se encontro comisión con este periodo", null);

            ReporteComisionVendedor_excel reporte = new ReporteComisionVendedor_excel();
            string rpt = reporte.GenerarReporte(listado,periodo);

            return new ResponseModel<string>(true, Constante.MESSSGE_SUCCESS_REPORT, rpt);
        }
    }
}
