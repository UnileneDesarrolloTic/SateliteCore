using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request.GestioEquipoEngaste;
using SatelliteCore.Api.Models.Response.GestioEquipoEngaste;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IGestionEquipoEngasteServices
    {
        public Task<IEnumerable<DatosFormatoEmpleado>> ObtenerEmpleado();
        public Task<IEnumerable<DatosFormatoListadoDadoEngaste>> ObtenerListadoDados();
        public Task<IEnumerable<DatosFormatoListarEquipoEngaste>> ListarEquipoEngaste(DatosFormularioFiltroEquipo dato);
        public Task<DatosFormatoInformacionEquipoEngaste> ObtenerInformacionEquipo(string idEquipo);
        public Task<ResponseModel<string>> RegistrarEquipoEngastado(DatosFormatoRegistroEquipoEngastado dato);
    }
}
