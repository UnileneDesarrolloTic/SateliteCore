using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;
using SatelliteCore.Api.ReportServices.Contracts.Dispensacion;
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

        public async Task<ResponseModel<string>> RegistrarDispensacionMP(DatosFormatoDispensacionMateriaPrima dato, string usuario)
        {
            if (dato.detalleDispensacion.Count == 0)
                return new ResponseModel<string>(false, "No hay información para registrar", "");


            dato.detalleDispensacion.ForEach(x =>
            {
                if (x.cantidadSolicitada < (x.cantidadDespachada + x.cantidadIngresada)) throw new ValidationModelException("El valor del ingreso excede a la cantidad solicitada");
                if (x.cantidadIngresada <= 0) throw new ValidationModelException("El valor deber ser mayor a 1");
            });


            string result = await _dispensacionRepository.RegistrarDispensacionMP(dato, usuario);
            return new ResponseModel<string>(true, "Registrado", "");
        }

        public async Task<IEnumerable<DatosFormatoHistorialDispensaccion>> HistorialDispensacionMP(string ordenFabricacion, string lote)
        {
            IEnumerable<DatosFormatoHistorialDispensaccion> resultado = new List<DatosFormatoHistorialDispensaccion>();
            resultado = await _dispensacionRepository.HistorialDispensacionMP(ordenFabricacion, lote);
            return resultado;
        }
        public async Task<ResponseModel<DatosFormatoInformacionDispensacionPT>> InformacionItem(string item, string ordenFabricacion, string secuencia)
        {
            DatosFormatoInformacionDispensacionPT lista = new DatosFormatoInformacionDispensacionPT();
            lista = await _dispensacionRepository.InformacionItem(item, ordenFabricacion, secuencia);
            return new ResponseModel<DatosFormatoInformacionDispensacionPT>(true, Constante.MESSAGE_SUCCESS, lista);
        }

        public async Task<DatosFormatoDispensacionDetalle> DetalleDispensacionReceta()
        {
            DatosFormatoDispensacionDetalle resultado = new DatosFormatoDispensacionDetalle();
            resultado = await _dispensacionRepository.DetalleDispensacionReceta();
            return resultado;
        }

        public async Task<ResponseModel<string>> RegistrarRecetasGlobal(List<DatosFormatoRegistroDispensacionRecetaGlobal> dato, string usuario)
        {
            if(dato.Count == 0)
               return new ResponseModel<string>(false, "No hay información para registrar", "");

            dato.ForEach(x =>
            {
                if (x.cantidadSolicitada < (x.cantidadDespachada + x.cantidadIngresada)) throw new ValidationModelException("El valor del ingreso excede a la cantidad solicitada");
                if (x.cantidadIngresada <= 0) throw new ValidationModelException("El valor deber ser mayor a 1");
            });

            IEnumerable<DatosFormatoRegistroDispensacionRecetaGlobal> registrado  =  new List<DatosFormatoRegistroDispensacionRecetaGlobal>();
            registrado = dato.Where(x => x.cantidadIngresada > 0);

            await _dispensacionRepository.RegistrarRecetasGlobal(registrado, usuario);
            return new ResponseModel<string>(true, "Registrado", "");
        }

        public async Task<IEnumerable<DatosFormatoDispensacionGuiaDespacho>> DispensacionGuiaDespacho(DatosFormatoFiltroDispensacion dato)
        {
            return await _dispensacionRepository.DispensacionGuiaDespacho(dato);
        }

        public async Task<IEnumerable<DatosFormatoMostrarDispensacionDespacho>> MostrarDispensacionDespacho(string id)
        {
            return await _dispensacionRepository.MostrarDispensacionDespacho(id);
        }

        public ResponseModel<string> GeneracionPdfDespacho(string id)
        {
            CodigoBarraGuiaDespacho_PDF claseReporte = new CodigoBarraGuiaDespacho_PDF();
            string reporte = claseReporte.Exportar(id);

            return new ResponseModel<string>(true,"generador de codigo barra", reporte);
        }
    }
}
