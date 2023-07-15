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
using System.Text;
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
            if (dato.agrupador.Count == 0)
                    throw new ValidationModelException("Los datos del agrupador no son válidos !!");

            IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion> listado = new List<DatosFormatoProgramacionOperacionesOrdenFabricacion>();

            StringBuilder unionAgrupador = new StringBuilder();

            foreach (int valor in dato.agrupador)
            {
                unionAgrupador.Append($"{valor},");
            }

            listado = await _programacionOperacionesRepository.ObtenerProgramacionOrdenFabricacion(dato, unionAgrupador.ToString());
            
            if (listado.Count() == 0)
                return new ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>>(false, "No hay información a mostrar", listado);

            return new ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>>(true, "Hay información a mostrar", listado);
        }

        public async Task<ResponseModel<string>> ActualizarFechaProgramada(DatosFormatoRegistrarFechaProgramacion dato, string usuario)
        {
            if (dato.fechaInicio == null && dato.fechaEntrega == null)
                return new ResponseModel<string>(false, "Debe ingresar la fecha de inicio o la fecha entrega", "");

            await _programacionOperacionesRepository.ActualizarFechaProgramada(dato, usuario);

            return new ResponseModel<string>(true, "Registrado", "");
        }

        public async Task<IEnumerable<DatosFormatoListadoFechaProgramadas>> ObtenerTipoFechaOrdenFabricacion(string ordenFabricacion, string tipoFecha)
        {   
            IEnumerable<DatosFormatoListadoFechaProgramadas> listado = new List<DatosFormatoListadoFechaProgramadas>();

            listado = await _programacionOperacionesRepository.ObtenerTipoFechaOrdenFabricacion(ordenFabricacion, tipoFecha);

            return listado;
        }

        public async Task<ResponseModel<string>> RegistrarDivisionProgramacion(DatosFormatoDividirRegistroProgramacion dato, string usuario)
        {
            int suma = 0;

            dato.divisionProgramacion.ForEach(x =>
            {
               if(x.cantidad < 1) throw new ValidationModelException("No puede ingresar cantidad menores a 1");
                suma = suma + x.cantidad;
            });

            if (suma != dato.cantidadProgramada)
                return new ResponseModel<string>(false, "La cantidad programada debe ser igual que la suma de la cantidad dividida", "");


            await _programacionOperacionesRepository.RegistrarDivisionProgramacion(dato, usuario);

            return new ResponseModel<string>(true, "Registrado", "");
        }
    }
}
