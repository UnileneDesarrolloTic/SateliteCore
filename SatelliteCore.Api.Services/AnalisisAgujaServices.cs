using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.AnalsisAguja;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class AnalisisAgujaServices : IAnalisisAgujaServices
    {
        private readonly IAnalisisAgujaRepository _analisisAgujaRepository;

        public AnalisisAgujaServices(IAnalisisAgujaRepository analisisAgujaRepository)
        {
            _analisisAgujaRepository = analisisAgujaRepository;
        }

        public async Task<IEnumerable<ListarAnalisisAgujaModel>> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina)
        {
            ordenCompra = string.IsNullOrEmpty(ordenCompra) ? null : ("%" + ordenCompra + "%");
            lote = string.IsNullOrEmpty(lote) ? null : ("%" + lote + "%");
            item = string.IsNullOrEmpty(item) ? null : ("%" + item + "%");

            return await _analisisAgujaRepository.ListarAnalisisAguja(ordenCompra, lote, item, pagina);
        }

        public async Task<IEnumerable<ListarOrdenCompra>> ListaOrdenesCompra(string numeroOrden)
        {
            if (string.IsNullOrEmpty(numeroOrden))
                numeroOrden = "";

            numeroOrden = "000000" + numeroOrden;

            if (numeroOrden.Substring(0, 3) != "FOR")
                numeroOrden = "FOR" + numeroOrden.Substring((numeroOrden.Length - 6), 6);

            return await _analisisAgujaRepository.ListaOrdenesCompra(numeroOrden);
        }

        public async Task<ResponseModel<int>> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia)
        {
            int cantidad = await _analisisAgujaRepository.CantidadPruebaFlexionPorItem(controlNumero, secuencia);
            return new ResponseModel<int>(true, Constante.MESSAGE_SUCCESS, cantidad);
        }

        public async Task<ResponseModel<object>> RegistrarAnalisisAguja(ControlAgujasModel matricula)
        {
            string cantidad = await _analisisAgujaRepository.RegistrarAnalisisAguja(matricula);

            return new ResponseModel<object>(true, Constante.MESSAGE_SUCCESS, new { numeroAnalisis = cantidad });
        }

        public async Task<ResponseModel<string>> ValidarLoteCreado(string controlNumero, int secuencia, int codUsuarioSesion)
        {
            await _analisisAgujaRepository.ValidarLoteCreado(controlNumero, secuencia, codUsuarioSesion);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, null);
        }

        public async Task<object> AnalisisAgujaFlexion(string loteAnalisis)
        {
            (ObtenerAnalisisAgujaModel cabecera, List<AnalisisAgujaFlexionEntity> detalle) = await _analisisAgujaRepository.AnalisisAgujaFlexion(loteAnalisis);

            if (string.IsNullOrEmpty(cabecera.ControlNumero))
                throw new NotFoundException("No se encontró el análisis");

            object result = new { cabecera, detalle };
            return result;
        }

        public async Task<ResponseModel<string>> GuardarEditarPruebaFlexionAguja(DatosFormatoRegistroPruebasAgujasModel dato)
        {          
            string loteAnalisis = dato.Lote;
            await _analisisAgujaRepository.EliminarPruebaFlexionAguja(loteAnalisis, dato.FechaAnalisis);

            ResponseModel<ObtenerDatosGeneralesDTO> datosGeneralesAnalisis = await ObtenerDatosGenerales(dato.Lote);

            if (datosGeneralesAnalisis.Content.Especialidad != dato.Especialidad)
            {
                await _analisisAgujaRepository.ActualizarEspecialidad(dato.Especialidad, dato.Lote);
                string serieActualizado = await _analisisAgujaRepository.ObtenerSeriePorLote(dato.Lote);

                if (datosGeneralesAnalisis.Content.Serie != serieActualizado)
                {
                    await _analisisAgujaRepository.ActualizarSerie(dato.Lote, serieActualizado);
                    await _analisisAgujaRepository.ActualizarCantidadPruebasFlexion(dato.Lote);
                }

            }

            if (dato.Especialidad == "N")
                await _analisisAgujaRepository.GuardarPruebaFlexionAguja(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos de la prueba de flexión !!");
        }


        public async Task<ResponseModel<string>> ReporteAnalisisFlexion(string loteAnalisis)
        {
            dynamic datosAnalisis = await AnalisisAgujaFlexion(loteAnalisis);

            if(datosAnalisis.cabecera.Especialidad == "S")
                return new ResponseModel<string>(false, "Aguja de especialidad no cuenta con prueba de flexión.", null);

            ObtenerAnalisisAgujaModel cabeceraAnalisis = datosAnalisis.cabecera;
            List<AnalisisAgujaFlexionEntity> detalleAnalisis = datosAnalisis.detalle;

            if (detalleAnalisis.Count < 1)
                return new ResponseModel<string>(false, "No cuenta con prueba de flexión", null);

            FlexionAguja flexionAguja = new FlexionAguja();
            string reporte = flexionAguja.GenerarReporte(loteAnalisis, cabeceraAnalisis, detalleAnalisis);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
        }

        public async Task<ResponseModel<ObtenerDatosGeneralesDTO>> ObtenerDatosGenerales(string loteAnalisis)
        {
            ObtenerDatosGeneralesDTO result = await _analisisAgujaRepository.ObtenerDatosGenerales(loteAnalisis);

            if (string.IsNullOrEmpty(result.ControlNumero))
                return new ResponseModel<ObtenerDatosGeneralesDTO>(false, "No se encontró el lote de análisis", result);
            
            return new ResponseModel<ObtenerDatosGeneralesDTO>(true, Constante.MESSAGE_SUCCESS, result);
        }

        public async Task<ResponseModel<AnalisisAgujaPlanMuestreoEntity>> ObtenerPlanMuestreo(string loteAnalisis)
        {
            AnalisisAgujaPlanMuestreoEntity result = await _analisisAgujaRepository.ObtenerPlanMuestreo(loteAnalisis);

            if (string.IsNullOrEmpty(result.LoteAnalisis))
                return new ResponseModel<AnalisisAgujaPlanMuestreoEntity>(false, "No se encontró el lote de análisis", result);

            return new ResponseModel<AnalisisAgujaPlanMuestreoEntity>(true, Constante.MESSAGE_SUCCESS, result);
        }

        public async Task<ResponseModel<string>> GuardarPlanMuestreo(AnalisisAgujaPlanMuestreoEntity planMuestreo)
        {
            planMuestreo.Fecha = DateTime.Now;

            await _analisisAgujaRepository.EliminarPlanMuestreo(planMuestreo.LoteAnalisis);
            await _analisisAgujaRepository.RegistrarPlanMuestreo(planMuestreo);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos del plan de muestreo !!");
        }

        public async Task<ResponseModel<List<AnalisisAgujaPruebaDimensionalEntity>>> ObtenerPruebaDimensional(string loteAnalisis)
        {
            List<AnalisisAgujaPruebaDimensionalEntity> result = await _analisisAgujaRepository.ObtenerPruebaDimensional(loteAnalisis);

            if (result.Count < 1)
                return new ResponseModel<List<AnalisisAgujaPruebaDimensionalEntity>>(false, "No se encontro el lote de análisis", result);

            return new ResponseModel<List<AnalisisAgujaPruebaDimensionalEntity>>(true, Constante.MESSAGE_SUCCESS, result);
        }

        public async Task<ResponseModel<string>> GuardarPruebaDimensional(List<AnalisisAgujaPruebaDimensionalEntity> prueba)
        {

            string loteAnalisis = prueba[0].LoteAnalisis;
            await _analisisAgujaRepository.EliminarPruebaDimensional(loteAnalisis);
            await _analisisAgujaRepository.RegistrarPruebaDimensional(prueba);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos de la prueba dimensional!!");
        }

        public async Task<ResponseModel<List<AnalisisAgujaElasticidadPerforacionEntity>>> ObtenerPruebaElasticidadPerforacion(string loteAnalisis)
        {
            List<AnalisisAgujaElasticidadPerforacionEntity> result = 
                (List<AnalisisAgujaElasticidadPerforacionEntity>) await _analisisAgujaRepository.ObtenerPruebaElasticidadPerforacion(loteAnalisis);

            if (result.Count < 1)
                return new ResponseModel<List<AnalisisAgujaElasticidadPerforacionEntity>>(false, "No se encontro el lote de análisis", result);

            return new ResponseModel<List<AnalisisAgujaElasticidadPerforacionEntity>>(true, Constante.MESSAGE_SUCCESS, result);
        }

        public async Task<ResponseModel<string>> GuardarPruebaElasticidadPerforacion(List<AnalisisAgujaElasticidadPerforacionEntity> prueba)
        {

            string loteAnalisis = prueba[0].LoteAnalisis;

            await _analisisAgujaRepository.EliminarPruebaElasticidadPerforacion(loteAnalisis);
            await _analisisAgujaRepository.RegistrarPruebaElasticidadPerforacion(prueba);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos de la prueba de elasticidad y perforación!!");
        }

        public async Task<ResponseModel<List<AnalisisAgujaPruebaAspectoEntity>>> ObtenerPruebaAspecto(string loteAnalisis)
        {
            List<AnalisisAgujaPruebaAspectoEntity> result = (List<AnalisisAgujaPruebaAspectoEntity>) await _analisisAgujaRepository.ObtenerPruebaAspecto(loteAnalisis);

            if (result.Count < 1)
                return new ResponseModel<List<AnalisisAgujaPruebaAspectoEntity>>(false, "No se encontro el lote de análisis", result);

            return new ResponseModel<List<AnalisisAgujaPruebaAspectoEntity>>(true, Constante.MESSAGE_SUCCESS, result);
        }

        public async Task<ResponseModel<string>> GuardarPruebaAspecto(PruebaAspectoYObservacionesDTO datos, int usuario)
        {
            if(string.IsNullOrEmpty(datos.Conclusion) || (datos.Conclusion != "R" && datos.Conclusion != "A" && datos.Conclusion != "S"))
                throw new ValidationModelException("Los datos enviados no son válidos.");

            DateTime fechaActual = DateTime.Now;

            List<AnalisisAgujaPruebaAspectoEntity> nuevoArreglo = datos.Pruebas.Select(item =>
            {
                if (!item.ValidarDatos())
                    throw new ValidationModelException("Los datos enviados no son válidos !!");

                item.Usuario = usuario;
                item.Fecha = fechaActual;

                return item;
            }).ToList();

            datos.Pruebas = nuevoArreglo;

            string loteAnalisis = datos.Pruebas[0].LoteAnalisis;

            await _analisisAgujaRepository.EliminarPruebaAspecto(loteAnalisis);
            await _analisisAgujaRepository.RegistrarPruebaAspecto(datos, loteAnalisis);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos de la prueba de aspecto!!");
        }

        public async Task<ResponseModel<string>> ObtenerReporteAnalisisAguja(string loteAnalisis)
        {

            Task<(ObtenerAnalisisAgujaModel _, List<AnalisisAgujaFlexionEntity> flexion)> analisisFlexion = _analisisAgujaRepository.AnalisisAgujaFlexion(loteAnalisis);
            Task<ObtenerDatosGeneralesDTO> datosGenerales = _analisisAgujaRepository.ObtenerDatosGenerales(loteAnalisis);
            Task<AnalisisAgujaPlanMuestreoEntity> planMuestreo = _analisisAgujaRepository.ObtenerPlanMuestreo(loteAnalisis);
            Task<List<AnalisisAgujaPruebaDimensionalEntity>> dimensionalCorrosion = _analisisAgujaRepository.ObtenerPruebaDimensional(loteAnalisis);
            Task<IEnumerable<AnalisisAgujaElasticidadPerforacionEntity>> elasticidadPerforacion = _analisisAgujaRepository.ObtenerPruebaElasticidadPerforacion(loteAnalisis);
            Task<IEnumerable<AnalisisAgujaPruebaAspectoEntity>> aspectoAguja = _analisisAgujaRepository.ObtenerPruebaAspecto(loteAnalisis);

            await Task.WhenAll(analisisFlexion, datosGenerales, planMuestreo, dimensionalCorrosion, elasticidadPerforacion, aspectoAguja);


            List<AnalisisAgujaFlexionEntity> flexionAux = analisisFlexion.Result.flexion.FindAll(x => x.TipoRegistro == 2).ToList();

            /* if (flexionAux.Count < 1)
                 return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No cuenta con prueba de flexión");*/

            if (string.IsNullOrEmpty(datosGenerales.Result.OrdenCompra))
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No cuenta con registros de análisis");

            if (string.IsNullOrEmpty(planMuestreo.Result.LoteAnalisis))
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No cuenta con prueba de funcionalidad y flexión");

            if (dimensionalCorrosion.Result.Count < 1)
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No cuenta con prueba de dimensión y corrosión");

            List<AnalisisAgujaElasticidadPerforacionEntity> elasticidadPerforacionAux = elasticidadPerforacion.Result.ToList();

            if (elasticidadPerforacionAux.Count < 1)
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No cuenta con prueba de elasticidad y perforación");

            List<AnalisisAgujaPruebaAspectoEntity> aspectoAgujaAux = aspectoAguja.Result.ToList();

            if (aspectoAgujaAux.Count < 1)
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No cuenta con prueba de aspecto de la aguja");

            PruebasAnalisis flexionAguja = new PruebasAnalisis();
            string reporte = flexionAguja.GenerarReporte(loteAnalisis, datosGenerales.Result, planMuestreo.Result, dimensionalCorrosion.Result, flexionAux, 
                elasticidadPerforacionAux, aspectoAgujaAux);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
        }

    }
}
