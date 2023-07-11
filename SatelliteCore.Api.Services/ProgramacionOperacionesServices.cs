using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.ProgramacionOperaciones;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.ProgramacionOperaciones;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class ProgramacionOperacionesServices : IProgramacionOperacionesServices
    {
        private readonly IProgramacionOperacionesRepository _programacionOperacionesRepository;

        public ProgramacionOperacionesServices(IProgramacionOperacionesRepository programacionOperacionesRepository)
        {
            _programacionOperacionesRepository = programacionOperacionesRepository;
          
        }

        public async Task<IEnumerable<DatosFormatoAgrupadores>> ObtenerAgrupadores(string gerencia)
        {
            IEnumerable<DatosFormatoAgrupadores> listado = new List<DatosFormatoAgrupadores>();
             listado = await _programacionOperacionesRepository.ObtenerAgrupadores(gerencia);

            return listado;
        }


            public async Task<ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>>> ObtenerProgramacionOrdenFabricacion(DatosFormatoProgramacionOperaciones dato)
        {
            if (dato.estado == "PR")
                if(dato.fechaInicio == null || dato.fechaFinal == null )
                    throw new ValidationModelException("Los datos enviados no son válidos !!");

            IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion> listado = new List<DatosFormatoProgramacionOperacionesOrdenFabricacion>();
            listado = await _programacionOperacionesRepository.ObtenerProgramacionOrdenFabricacion(dato);
            
            if (listado.Count() == 0)
                return new ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>>(false, "No hay información a mostrar", listado);

            return new ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>>(true, "Hay información a mostrar", listado);
        }

        public async Task<ResponseModel<string>> ActualizarFechaProgramada(DatosFormatoRegistrarFechaProgramacion dato, string usuario)
        {
           await _programacionOperacionesRepository.ActualizarFechaProgramada(dato, usuario);

            return new ResponseModel<string>(true, "Registro Exitoso", "");
        }
    }
}
