using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Request.ProgramacionOperaciones;
using SatelliteCore.Api.Models.Response.ProgramacionOperaciones;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IProgramacionOperacionesRepository
    {
        public Task<IEnumerable<DatosFormatoAgrupadores>> ObtenerAgrupadores(string gerencia);
        public Task<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>> ObtenerProgramacionOrdenFabricacion(DatosFormatoProgramacionOperaciones dato, string unionAgrupador);
        public Task<string> ActualizarFechaProgramada(DatosFormatoRegistrarFechaProgramacion dato, string usuario);
        public Task<IEnumerable<DatosFormatoListadoFechaProgramadas>> ObtenerTipoFechaOrdenFabricacion(string ordenFabricacion, string tipoFecha);
        public Task<string> RegistrarDivisionProgramacion(DatosFormatoDividirRegistroProgramacion dato, string usuario);
    }
}
