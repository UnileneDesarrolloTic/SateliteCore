using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Request.ProgramacionOperaciones;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.ProgramacionOperaciones;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IProgramacionOperacionesServices
    {
        public Task<IEnumerable<DatosFormatoAgrupadores>> ObtenerAgrupadores(string gerencia);
        public Task<ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>>> ObtenerProgramacionOrdenFabricacion(DatosFormatoProgramacionOperaciones dato);
        public Task<ResponseModel<string>> ActualizarFechaProgramada(DatosFormatoRegistrarFechaProgramacion dato, string usuario);
    }
}
