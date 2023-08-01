using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.GestioEquipoEngaste;
using SatelliteCore.Api.Models.Request.GestioEquipoEngaste;
using SatelliteCore.Api.Services.Contracts;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class GestionEquipoEngasteServices : IGestionEquipoEngasteServices
    {
        private readonly IGestionEquipoEngasteRepository _gestionEquipoEngasteRepository;

        public GestionEquipoEngasteServices(IGestionEquipoEngasteRepository gestionEquipoEngasteRepository)
        {
            _gestionEquipoEngasteRepository = gestionEquipoEngasteRepository;

        }

        public async Task<IEnumerable<DatosFormatoEmpleado>> ObtenerEmpleado()
        {
            return await _gestionEquipoEngasteRepository.ObtenerEmpleado();
        }

        public async Task<IEnumerable<DatosFormatoListadoDadoEngaste>> ObtenerListadoDados()
        {
            return await _gestionEquipoEngasteRepository.ObtenerListadoDados();
        }

        public async Task<IEnumerable<DatosFormatoListarEquipoEngaste>> ListarEquipoEngaste(DatosFormularioFiltroEquipo dato)
        {
            return await _gestionEquipoEngasteRepository.ListarEquipoEngaste(dato);
        }

        public async Task<DatosFormatoInformacionEquipoEngaste> ObtenerInformacionEquipo(string idEquipo)
        {   
            if(string.IsNullOrEmpty(idEquipo))
                    throw new ValidationModelException("verificar los parametros enviados");

            return await _gestionEquipoEngasteRepository.ObtenerInformacionEquipo(idEquipo);
        }

        public async Task<ResponseModel<string>> RegistrarEquipoEngastado(DatosFormatoRegistroEquipoEngastado dato)
        {
            string resultado = "";
            if (string.IsNullOrEmpty(dato.nombre) || string.IsNullOrEmpty(dato.Tipo) || dato.idpersona == 0)
                throw new ValidationModelException("verificar los parametros enviados");

            resultado = await _gestionEquipoEngasteRepository.RegistrarEquipoEngastado(dato);

            return  new ResponseModel<string>(true,"Registrado","");
        }

    }
}
