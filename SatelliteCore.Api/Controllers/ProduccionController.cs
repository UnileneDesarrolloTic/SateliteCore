using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.OCDrogueria;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ProduccionController : ControllerBase
    {
        private readonly IProduccionServices _pronosticoServices;

        public ProduccionController(IProduccionServices pronosticoServices)
        {
            _pronosticoServices = pronosticoServices;
        }

        [HttpGet("ProductosArima")]
        public async Task<ActionResult> SeguimientoProductosArima(string periodo)
        {
            List<ProductoArimaModel> listaCandidatos = await _pronosticoServices.SeguimientoProductosArima(periodo);

            return Ok(listaCandidatos);
        }

        [HttpPost("ListaPedidosCreadoAuto")]
        public async Task<ActionResult> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro)
        {
            (IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros) = await _pronosticoServices.ListaPedidosCreadoAuto(filtro);

            PaginacionModel<PedidosCreadosAutoLogModel> PedidosPaginados =
                new PaginacionModel<PedidosCreadosAutoLogModel>((List<PedidosCreadosAutoLogModel>)ListaPedidos, filtro.Pagina, filtro.RegistrosPorPagina, TotalRegistros);

            return Ok(PedidosPaginados);
        }

        [HttpGet("SegimientoCandidatosMP")]
        public async Task<ActionResult> SegimientoCandidatosMP(string regla)
        {
            SeguimientoCandMPAGenericModel listaCandidatos = await _pronosticoServices.ListaSeguimientoCandidatosMP(regla);
            return Ok(listaCandidatos);
        }

        [HttpGet("ExportarAgujasMateriaPrima")]
        public async Task<ActionResult> ExportarAgujasMateriaPrima(string regla)
        {
            ResponseModel<string> respuesta = await _pronosticoServices.ExportarAgujasMateriaPrima(regla);
            return Ok(respuesta);
        }




        [HttpGet("ControlCalidadItemMP")]
        public async Task<ActionResult> ControlCalidadItemMP(string Item)
        {
            List<DetalleControlCalidadItemMP> listaDetalleControlCalidad = await _pronosticoServices.ControlCalidadItemMP(Item);
            return Ok(listaDetalleControlCalidad);
        }

        [HttpGet("MostrarColumnaMP")]
        public async Task<ActionResult> MostrarColumnaMP()
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                             new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            bool Permiso = await _pronosticoServices.MostrarColumnaMP(idUsuario);
            ResponseModel<dynamic> responseSuccesss = new ResponseModel<dynamic>(true, Constante.MESSAGE_SUCCESS, new { permisoColumna = Permiso });

            return Ok(responseSuccesss);
        }


        [HttpPost("CompraMateriaPrima")]
        public async Task<ActionResult> PronosticoCompraMP(PronosticoCompraMP dato)
        {

            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                             new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            List<CompraMPArimaModel> listaProductos = await _pronosticoServices.SeguimientoCompraMPArima(dato);

            return Ok(listaProductos);
        }

        [HttpPost("CompraMateriaPrimaExportar")]
        public async Task<ActionResult> CompraMateriaPrimaExportar(PronosticoCompraMP dato)
        {

            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            ResponseModel<string> respuesta = await _pronosticoServices.CompraMateriaPrimaExportar(dato);

            return Ok(respuesta);
        }

        [HttpGet("LoteFabricacionEtiquetas")]
        public async Task<ActionResult> LoteFabricacionEtiquetas(string NumeroLote)
        {

            if (NumeroLote == "")
            {
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }
            ResponseModel<FormatoEstructuraLoteEtiquetas> response = await _pronosticoServices.LoteFabricacionEtiquetas(NumeroLote);
            return Ok(response);
        }


        [HttpPost("RegistrarLoteFabricacionEtiquetas")]
        public async Task<ActionResult> RegistrarLoteFabricacionEtiquetas(List<DatosEstructuraLoteEtiquetasModel> dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response = await _pronosticoServices.RegistrarLoteFabricacionEtiquetas(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("ListarLoteEstado")]
        public async Task<ActionResult> ListarLoteEstado()
        {
            IEnumerable<DatoFormatoLoteEstado> response = await _pronosticoServices.ListarLoteEstado();
            return Ok(response);
        }

        [HttpPost("ModificarLoteEstado")]
        public async Task<ActionResult> ModificarLoteEstado(DatosFormatoRequestLoteEstado dato)
        {
            ResponseModel<string> response = await _pronosticoServices.ModificarLoteEstado(dato);
            return Ok(response);
        }

        [HttpGet("ListarItemOrdenCompra")]
        public async Task<ActionResult> ListarItemOrdenCompra(string Anio)
        {
            DatosFormatoInformacionCalendarioSeguimientoOC response = await _pronosticoServices.ListarItemOrdenCompra(Anio);
            return Ok(response);
        }

        [HttpGet("BuscarItemOrdenCompra")]
        public async Task<ActionResult> BuscarItemOrdenCompra(string Item, string Anio)
        {
            DatosFormatoInformacionItemOrdenCompra response = await _pronosticoServices.BuscarItemOrdenCompra(Item, Anio);
            return Ok(response);
        }


        [HttpPost("ActualizarFechaPrometida")]
        public async Task<ActionResult> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra datos)
        {
            ResponseModel<string> response = await _pronosticoServices.ActualizarFechaPrometida(datos);
            return Ok(response);
        }

        [HttpGet("VisualizarOrdenCompra")]
        public async Task<ActionResult> VisualizarOrdenCompra(string OrdenCompra)
        {
            (object cabecera, object detalle) = await _pronosticoServices.VisualizarOrdenCompra(OrdenCompra);
            object response = new { cabecera, detalle };

            return Ok(response);
        }

        [HttpPost("ActualizarFechaPrometidaMasiva")]
        public async Task<ActionResult> ActualizarFechaPrometidaMasiva(DatosFormatoCabeceraOrdenCompraModel datos)
        {
            ResponseModel<string> response = await _pronosticoServices.ActualizarFechaPrometidaMasiva(datos);
            return Ok(response);
        }


        [HttpGet("SeguimientoOCDrogueria")]
        public async Task<ActionResult> SeguimientoOCDrogueria(int idproveedor)
        {
            if (string.IsNullOrEmpty(idproveedor.ToString()))
                throw new ValidationModelException("El proveedor es obligatorio");

            ResponseModel<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>> response = await _pronosticoServices.SeguimientoOCDrogueria(idproveedor);
            return Ok(response);
        }


        [HttpGet("MostrarOrdenCompraDrogueria")]
        public async Task<ActionResult> MostrarOrdenCompraDrogueria(string Item)
        {
            ResponseModel<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>> response = await _pronosticoServices.MostrarOrdenCompraDrogueria(Item);
            return Ok(response);
        }
            
        [HttpGet("MostrarProveedorDrogueria")]
        public async Task<ActionResult> MostrarProveedorDrogueria ()
        {
            ResponseModel<IEnumerable<DatosFormatoMostrarProveedorDrogueria>> response = await _pronosticoServices.MostrarProveedorDrogueria();
            return Ok(response);
        }


        [HttpGet("ExcelCompraDrogueria")]
        public async Task<ActionResult> ExcelCompraDrogueria(int idproveedor, bool mostrarcolumna, string agrupador)
        {
            if (string.IsNullOrEmpty(idproveedor.ToString()))
                throw new ValidationModelException("El proveedor es obligatorio");

            ResponseModel<string> response = await _pronosticoServices.ExcelCompraDrogueria(idproveedor,mostrarcolumna, agrupador);
            return Ok(response);
        }

        [HttpGet("MostrarOrdenCompraPrevios")]
        public async Task<ActionResult> MostrarOrdenCompraPrevios()
        {
            ResponseModel<IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio>> response = await _pronosticoServices.MostrarOrdenCompraPrevios();
            return Ok(response);
        }

        [HttpGet("VisualizarOrdenCompraSimulada")]
        public async Task<ActionResult> VisualizarOrdenCompraSimulada(string proveedor)
        {

            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);

            (object cabecera, object detalle) = await _pronosticoServices.VisualizarOrdenCompraSimulada(proveedor, usuario);
            object response = new { cabecera, detalle };

            return Ok(response);
        }

        [HttpPost("GuardarOrdenCompraVencida")]
        public async Task<ActionResult> GuardarOrdenCompraVencida(DatosFormatoCambiarEstadoOCVencida dato)
        {   
              if(string.IsNullOrEmpty(dato.item) || string.IsNullOrEmpty(dato.numeroOrden))
                    throw new ValidationModelException("verificar los parametros enviados");

            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> respuesta = await _pronosticoServices.GuardarOrdenCompraVencida(dato, usuario);
            return Ok(respuesta);
        }

        [HttpGet("GenerarOrdenCompraDrogueria")]
        public async Task<ActionResult> GenerarOrdenCompraDrogueria()
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> respuesta = await _pronosticoServices.GenerarOrdenCompraDrogueria(idUsuario);
            return Ok(respuesta);
        }

        [HttpPost("RegistrarOrdenCompraDrogueria")]
        public async Task<ActionResult> RegistrarOrdenCompraDrogueria(DatosFormatoGuardarCabeceraOrdenCompraDrogueria dato)
        {
            int encontrarPendiente = dato.detalle.FindIndex(x => x.estado == "PE"); 
            if(encontrarPendiente != -1)
                throw new ValidationModelException("verificar los parametros enviados");
           
            string strusuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            int idusuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> respuesta = await _pronosticoServices.RegistrarOrdenCompraDrogueria(dato, strusuario, idusuario);
            return Ok(respuesta);
        }

        [HttpGet("InformacionSeguimientoAguja")]
        public async Task<ActionResult> InformacionSeguimientoAguja()
        {
            DatosInformacionGeneralReporteCompraArimaAgujas respuesta = await _pronosticoServices.InformacionSeguimientoAguja();
            return Ok(respuesta);
        }

        [HttpGet("InformacionSeguimientoAgujaExcel")]
        public async Task<ActionResult> InformacionSeguimientoAgujaExcel(string mostrarColumna)
        {
            if (string.IsNullOrEmpty(mostrarColumna))
                throw new ValidationModelException("verificar los parametros enviados");

            ResponseModel<string> respuesta = await _pronosticoServices.InformacionSeguimientoAgujaExcel(mostrarColumna);
            return Ok(respuesta);
        }

        [HttpGet("MostrarOrdenCompraArima")]
         public async Task<ActionResult> MostrarOrdenCompraArima (string Item, string Tipo)
        {
            ResponseModel<IEnumerable<DatosFormatoTransitoPendienteOC>> respuesta = await _pronosticoServices.MostrarOrdenCompraArima(Item, Tipo);
            return Ok(respuesta);
        }
    }
}


