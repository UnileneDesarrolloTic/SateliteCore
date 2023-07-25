using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Response.GestioEquipoEngaste;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IGestionEquipoEngasteRepository
    {
        public Task<IEnumerable<DatosFormatoEmpleado>> ObtenerEmpleado();
        public Task<IEnumerable<DatosFormatoListadoDadoEngaste>> ObtenerListadoDados();
        public Task<IEnumerable<DatosFormatoListarEquipoEngaste>> ListarEquipoEngaste(DatosFormularioFiltroEquipo dato);
        public Task<DatosFormatoInformacionEquipoEngaste> ObtenerInformacionEquipo(string idEquipo);
        public Task<string> RegistrarEquipoEngastado(DatosFormatoRegistroEquipoEngastado dato);
    }
}
