﻿using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.RRHH;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IRRHHServices
    {
        public Task<ResponseModel<IEnumerable<ReporteAsistenciaDTO>>> ListarAsistencia(DateTime fecha, int usuarioToken);
        public Task<ResponseModel<string>> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera dato, string usuario);
        public Task<IEnumerable<DatosFormatoListarHorasExtrasPersona>> ListarHoraExtrasPersona(DatosFormatoFiltraHoraExtras dato);
        public Task<object> BuscarInformacionHorasExtrasPersona(int Cabecera);
        public Task ProcesarHorasExtrasPlanilla(string periodo);
        public Task<string> ReporteHorasExtrasGeneradas(string periodo);
        public Task<string> FormatoAutorizacionSobretiempo(int id);
        public Task<ResponseModel<string>> AutorizacionSobretiempoPorPersona(string periodo);
        public Task<IEnumerable<DatosFormatoReporteComisionVendedor>> ReporteComisionVendedor(string periodo);
        public Task<ResponseModel<string>> ReporteComisionVendedorExcel(string periodo);
    }
}
