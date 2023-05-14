using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Report.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.RRHH;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.RRHH;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IRRHHRepository
    {
        public Task<IEnumerable<ReporteAsistenciaDTO>> ListarAsistencia(DateTime fecha, int usuarioToken);
        public Task<int> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera data, string usuario);
        public Task<IEnumerable<DatosFormatoListarHorasExtrasPersona>> ListarHoraExtrasPersona(DatosFormatoFiltraHoraExtras data);
        public Task<(DatosFormatoHorasExtrasCabeceraModel, List<DatosFormatoHorasExtrasDetalle>)> BuscarInformacionHorasExtrasPersona(int Cabecera);
        public Task ProcesarHorasExtrasPlanilla(string periodo);
        public Task<List<HorasExtraExportDTO>> ListarHorasExtrasGeneradas(string periodo);
        public Task<AutorizacionSobretiempoPersonaDTO> ListarHorasExtraExportas(string periodo);
        public Task<IEnumerable<DatosFormatoReporteComisionVendedor>> ReporteComisionVendedor(string periodo);

    }
}
