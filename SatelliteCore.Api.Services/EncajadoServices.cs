using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Encajado;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.Encajado;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class EncajadoServices : IEncajadoServices
    {
        private readonly IEncajadoRespository _encajadoRespository;

        public EncajadoServices(IEncajadoRespository encajadoRespository)
        {
            _encajadoRespository = encajadoRespository;
        }

        public async Task<ResponseModel<List<ListaOrdenesFabricaciónDTO>>> ListaOrdenesFabricacion(string ordenFabricacion, string lote)
        {

            List<ListaOrdenesFabricaciónDTO> lista = await _encajadoRespository.ListaOrdenesFabricacion(ordenFabricacion, lote);

            if (lista.Count < 1)
                return new ResponseModel<List<ListaOrdenesFabricaciónDTO>>(false, "No se han encontrado registros.", lista);

            return new ResponseModel<List<ListaOrdenesFabricaciónDTO>>(lista);
        }

        public async Task<ResponseModel<List<TransferenciaEncajadoDTO>>> ListaTransferenciasEncaje(string ordenFabricacion)
        {
            if (string.IsNullOrWhiteSpace(ordenFabricacion))
                throw new ValidationModelException();

            List<TransferenciaEncajadoDTO> lista = await _encajadoRespository.ListaTransferenciasEncaje(ordenFabricacion);

            if (lista.Count < 1)
                return new ResponseModel<List<TransferenciaEncajadoDTO>>(false, "No se han encontrado registros.", lista);

            return new ResponseModel<List<TransferenciaEncajadoDTO>>(lista);
        }

        public async Task<ResponseModel<string>> RegistarNuevaTrasnferencia(string ordenFabricacion, decimal cantidad, string usuario)
        {
            if (string.IsNullOrWhiteSpace(ordenFabricacion) || cantidad < 1 || string.IsNullOrWhiteSpace(usuario))
                throw new ValidationModelException();

            int registro = await _encajadoRespository.RegistarNuevaTrasnferencia(ordenFabricacion, cantidad, usuario);

            if (registro > 0)
                return new ResponseModel<string>(null);

            return new ResponseModel<string>(false, "Error al registrar la transferencia", "");
        }

        public async Task<ResponseModel<object>> ListraAsignacionesEncajePorEtapa(int idEncaje, int etapa)
        {
            if (idEncaje < 1 || etapa > 4 || etapa < 1)
                throw new ValidationModelException();

            (decimal totalAnterior, decimal totalActual, List<AsignacionEncajadoDTO>lista) result = await _encajadoRespository.ListraAsignacionesEncajePorEtapa(idEncaje, etapa);

            object response = new
            {
                result.totalAnterior,
                result.totalActual,
                result.lista
            };

            return new ResponseModel<object>(response);
        }

        public async Task<ResponseModel<string>> RegistrarAsignacion(DatosRegistrarAsignacionDTO asignacion)
        {
            if (!asignacion.ValidarDatos())
                throw new ValidationModelException();

            int registro = await _encajadoRespository.RegistrarAsignacionEncaje(asignacion);

            if(registro == -1)
                return new ResponseModel<string>(false, "El código del trabajador no existe, o no esta activo.", null);

            if (registro == 0)
                return new ResponseModel<string>(false, "No se pudo registrar la asignación.", null);

            return new ResponseModel<string>(null);
        }

        public async Task<ResponseModel<string>> ActualizaEstadoAsignacion(int id, string estado, string usuario)
        {
            if(id < 1 || string.IsNullOrWhiteSpace(estado) || string.IsNullOrWhiteSpace(usuario))
                throw new ValidationModelException();

            await _encajadoRespository.ActualizaEstadoAsignacion(id, estado, usuario);

            return new ResponseModel<string>(null);
        }
        public async Task<ResponseModel<string>> ReporteAsignacion(DateTime fechaInicio, DateTime fechaFin)
        {
            if(!Shared.ValidarFecha(fechaInicio) || !Shared.ValidarFecha(fechaFin))
                throw new ValidationModelException();

            List<DatosReporteEncajadoDTO> datosReporte = await _encajadoRespository.DatosReporteAsignacion(fechaInicio, fechaFin);

            if(datosReporte.Count < 1)
                return new ResponseModel<string>(false, "No se ha encontrado registros...", null);

            ReporteEncajado_Excel reporte = new ReporteEncajado_Excel();
            string reporteBase64 = reporte.GenerarReporte(datosReporte);

            if (string.IsNullOrWhiteSpace(reporteBase64))
                throw new Exception("Error al generar el reporte.");

            return new ResponseModel<string>(reporteBase64);


        }
    }
}
