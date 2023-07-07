using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class DispensacionServices : IDispensacionServices
    {
        private readonly IDispensacionRepository _dispensacionRepository;

        public DispensacionServices(IDispensacionRepository dispensacionRepository)
        {
            _dispensacionRepository = dispensacionRepository;
           
        }

        public async Task<ResponseModel<IEnumerable<DatosFormatoObtenerOrdenFabricacion>>> ObtenerOrdenFabricacion(DatosFormatoFiltroOrdenFabricacion dato)
        {
            if (string.IsNullOrEmpty(dato.estado))
                throw new Exception("Error de parametro");

            IEnumerable<DatosFormatoObtenerOrdenFabricacion> listado = new List<DatosFormatoObtenerOrdenFabricacion>();
            listado = await _dispensacionRepository.ObtenerOrdenFabricacion(dato);
            if (listado.Count() == 0)
                return new ResponseModel<IEnumerable<DatosFormatoObtenerOrdenFabricacion>>(false, "No hay información a mostrar", listado);

            return new ResponseModel<IEnumerable<DatosFormatoObtenerOrdenFabricacion>>(true, "Hay información a mostrar", listado);
        }
        public async Task<IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion>> RecetasOrdenFabricacion(string ordenFabricacion)
        {
            return await _dispensacionRepository.RecetasOrdenFabricacion(ordenFabricacion);
        }

        public async Task<ResponseModel<string>> RegistrarDispensacionMP(List<DatosFormatoDispensacionDetalleMP> dato, string usuario)
        {
            string result = await _dispensacionRepository.RegistrarDispensacionMP(dato, usuario);
            return new ResponseModel<string>(true, "Registrado", "");
        }

    }
}
