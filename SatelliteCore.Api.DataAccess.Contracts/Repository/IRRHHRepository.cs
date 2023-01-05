using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
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
    }
}
