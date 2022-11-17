using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class AnalisisAgujaController : ControllerBase
    {
        private readonly IAnalisisAgujaServices _analisisAgujaServices;

        public AnalisisAgujaController(IAnalisisAgujaServices analisisAgujaServices)
        {
            _analisisAgujaServices = analisisAgujaServices;
        }

        [HttpGet("ListarAnalisisAguja")]
        public async Task<IActionResult> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina)
        {
            if (pagina == 0)
                throw new ValidationModelException("El número de página es incorrecto");

            IEnumerable<ListarAnalisisAgujaModel> listaAnalisis = await _analisisAgujaServices.ListarAnalisisAguja(ordenCompra, lote, item, pagina);

            return Ok(listaAnalisis);
        }

        [HttpGet("ListaOrdenesCompra")]
        public async Task<IActionResult> ListaOrdenesCompra(string numeroOrden)
        {
            if (string.IsNullOrEmpty(numeroOrden))
                throw new ValidationModelException("El código FOR es obligatorio");

            IEnumerable<ListarOrdenCompra> listaOrdenesCompra = await _analisisAgujaServices.ListaOrdenesCompra(numeroOrden);
            return Ok(listaOrdenesCompra);
        }

        [HttpGet("CantidadPruebasFlexionPorItem")]
        public async Task<IActionResult> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia)
        {
            ResponseModel<int> cantidadPruebas = await _analisisAgujaServices.CantidadPruebaFlexionPorItem(controlNumero, secuencia);
            return Ok(cantidadPruebas);
        }

        [HttpPost("RegistrarAnalisisAguja")]
        public async Task<IActionResult> RegistrarAnalisisAguja(ControlAgujasModel matricula)
        {
            if (!ModelState.IsValid)
            {
                List<string> listaErrores = ModelState.Values.SelectMany(m => m.Errors).Select(e => e.ErrorMessage).ToList();
                throw new ValidationModelException(listaErrores);
            }

            matricula.CodUsuarioSesion = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<object> cantidadPruebas = await _analisisAgujaServices.RegistrarAnalisisAguja(matricula);
            return Ok(cantidadPruebas);
        }

        [HttpGet("ValidarLoteCreado")]
        public async Task<IActionResult> ValidarLoteCreado(string controlNumero, int secuencia)
        {
            if (string.IsNullOrEmpty(controlNumero) || secuencia == 0)
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            int codUsuarioSesion = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> cantidadPruebas = await _analisisAgujaServices.ValidarLoteCreado(controlNumero, secuencia, codUsuarioSesion);

            return Ok(cantidadPruebas);
        }

        [HttpGet("AnalisisAgujaFlexion")]
        public async Task<IActionResult> AnalisisAgujaFlexion(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            object analisis = await _analisisAgujaServices.AnalisisAgujaFlexion(loteAnalisis);

            return Ok(analisis);
        }

        [HttpPost("GuardarEditarPruebaFlexionAguja")]
        public async Task<IActionResult> GuardarEditarPruebaFlexion(DatosFormatoRegistroPruebasAgujasModel dato)
        {
            if (dato.Especialidad == "N")
            {
                if (!ModelState.IsValid) 
                    throw new ValidationModelException("Los datos de prueba de flexion no son válidos !!");
            }
               
            ResponseModel<string> result = await _analisisAgujaServices.GuardarEditarPruebaFlexionAguja(dato);

            return Ok(result);
        }

        [HttpGet("ReporteAnalisisFlexion")]
        public async Task<IActionResult> ReporteAnalisisFlexion(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            ResponseModel<string> reporte = await _analisisAgujaServices.ReporteAnalisisFlexion(loteAnalisis);

            return Ok(reporte);
        }

        [HttpGet("ObtenerDatosGenerales")]
        public async Task<IActionResult> ObtenerDatosGenerales(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            ResponseModel<ObtenerDatosGeneralesDTO> reporte = await _analisisAgujaServices.ObtenerDatosGenerales(loteAnalisis);

            return Ok(reporte);
        }

        [HttpGet("ObtenerPlanMuestreo")]
        public async Task<IActionResult> ObtenerPlanMuestreo(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            ResponseModel<AnalisisAgujaPlanMuestreoEntity> reporte = await _analisisAgujaServices.ObtenerPlanMuestreo(loteAnalisis);

            return Ok(reporte);
        }

        [HttpPost("GuardarPlanMuestreo")]
        public async Task<IActionResult> GuardarPlanMuestreo(AnalisisAgujaPlanMuestreoEntity planMuestreo)
        {
            if (!planMuestreo.ValidarDatos())
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            planMuestreo.Usuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> reporte = await _analisisAgujaServices.GuardarPlanMuestreo(planMuestreo);

            return Ok(reporte);
        }

        [HttpGet("ObtenerPruebaDimensional")]
        public async Task<IActionResult> ObtenerPruebaDimensional(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            ResponseModel<List<AnalisisAgujaPruebaDimensionalEntity>> result = await _analisisAgujaServices.ObtenerPruebaDimensional(loteAnalisis);

            return Ok(result);
        }

        [HttpPost("GuardarPruebaDimensional")]
        public async Task<IActionResult> GuardarPruebaDimensional(List<AnalisisAgujaPruebaDimensionalEntity> prueba)
        {
            int usuarioToken = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            DateTime fechaActual = DateTime.Now;

            List<AnalisisAgujaPruebaDimensionalEntity> nuevoArreglo = prueba.Select( item =>
            {
                if(!item.ValidarDatos())
                    throw new ValidationModelException("Los datos enviados no son válidos !!");
                
                item.Usuario = usuarioToken;
                item.Fecha = fechaActual;

                return item;
            }).ToList();
                

            ResponseModel<string> reporte = await _analisisAgujaServices.GuardarPruebaDimensional(nuevoArreglo);

            return Ok(reporte);
        }

        [HttpGet("ObtenerPruebaElasticidadPerforacion")]
        public async Task<IActionResult> ObtenerPruebaElasticidadPerforacion(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            ResponseModel<List<AnalisisAgujaElasticidadPerforacionEntity>> result = await _analisisAgujaServices.ObtenerPruebaElasticidadPerforacion(loteAnalisis);

            return Ok(result);
        }

        [HttpPost("GuardarPruebaElasticidadPerforacion")]
        public async Task<IActionResult> GuardarPruebaElasticidadPerforacion(List<AnalisisAgujaElasticidadPerforacionEntity> prueba)
        {
            int usuarioToken = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            DateTime fechaActual = DateTime.Now;

            List<AnalisisAgujaElasticidadPerforacionEntity> nuevoArreglo = prueba.Select(item =>
            {
                if (!item.DatosValidos())
                    throw new ValidationModelException("Los datos enviados no son válidos !!");

                item.Usuario = usuarioToken;
                item.Fecha = fechaActual;

                return item;
            }).ToList();

            ResponseModel<string> reporte = await _analisisAgujaServices.GuardarPruebaElasticidadPerforacion(nuevoArreglo);

            return Ok(reporte);
        }

        [HttpGet("ObtenerPruebaAspecto")]
        public async Task<IActionResult> ObtenerPruebaAspecto(string loteAnalisis)
        {
            if (string.IsNullOrEmpty(loteAnalisis))
                throw new ValidationModelException("Los datos enviados no son válidos !!");

            ResponseModel<List<AnalisisAgujaPruebaAspectoEntity>> result = await _analisisAgujaServices.ObtenerPruebaAspecto(loteAnalisis);

            return Ok(result);
        }


        [HttpPost("GuardarPruebaAspecto")]
        public async Task<IActionResult> GuardarPruebaAspecto(PruebaAspectoYObservacionesDTO datos)
        {
            int usuarioToken = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> reporte = await _analisisAgujaServices.GuardarPruebaAspecto(datos, usuarioToken);

            return Ok(reporte);
        }


        [HttpGet("ReporteAnalisisAguja")]
        public async Task<IActionResult> GuardarPruebaAspecto(string loteAnalisis)
        {
            ResponseModel<string> reporte = await _analisisAgujaServices.ObtenerReporteAnalisisAguja(loteAnalisis);

            return Ok(reporte);
        }

    }
}
