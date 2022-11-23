using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request.GestionCalidad;
using SatelliteCore.Api.Models.Request.GestorDocumentario;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.GestionCalidad;
using SatelliteCore.Api.ReportServices.Contracts.GestionCalidad;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class GestionCalidadServices : IGestionCalidadServices
    {
        private readonly IGestionCalidadRepository _gestionCalidadRepository;
        private readonly ICommonRepository _commonRepository;

        public GestionCalidadServices(IGestionCalidadRepository gestionCalidadRepository, ICommonRepository commonRepository)
        {
            _gestionCalidadRepository = gestionCalidadRepository;
            _commonRepository = commonRepository;
        }

        public async Task<List<MateriaPrimaDTO>> ObtenerMateriaPrima(string tipoConsulta, string lote)
        {
            if (string.IsNullOrEmpty(tipoConsulta) || string.IsNullOrEmpty(lote))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            if (tipoConsulta != "PT" && tipoConsulta != "MP")
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            return await _gestionCalidadRepository.ObtenerMateriaPrima(tipoConsulta, lote);
        }

        public async Task<DetalleSeguimientoLoteDTO> DetalleSeguimientoPorLote(RequestLotesDetalleDTO filtros)
        {

            if (filtros.Lotes.Count < 1 || filtros.OrdenesFabricacion.Count < 1)
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            DetalleSeguimientoLoteDTO detalle = new DetalleSeguimientoLoteDTO();

            detalle.OrdenesDeCompra = await _gestionCalidadRepository.OrdenCompraPorlote(filtros.Lotes);
            detalle.OrdenesDeFabricacion = await _gestionCalidadRepository.OrdenFabricacionPorlotes(filtros.OrdenesFabricacion);
            detalle.DocumentosPedidos = await _gestionCalidadRepository.OrdenDocumentosPedidosPorLotes(filtros.OrdenesFabricacion);
            detalle.GuiasRelacionadas = await _gestionCalidadRepository.OrdenGuiasRelacionadasPorLotes(filtros.OrdenesFabricacion);

            return detalle;
        }


        public async Task<List<VentasPorClienteDTO>> VentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            if (!filtros.ValidarDatos() || !Shared.ValidarFecha(filtros.FechaInicio) || !Shared.ValidarFecha(filtros.FechaFin) || filtros.FechaInicio > filtros.FechaFin)
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            return await _gestionCalidadRepository.VentasPorCliente(filtros);
        }

        public async Task<ResponseModel<string>> ReporteVentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            if (!filtros.ValidarDatos() || !Shared.ValidarFecha(filtros.FechaInicio) || !Shared.ValidarFecha(filtros.FechaFin) || filtros.FechaInicio > filtros.FechaFin)
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            List<VentasPorClienteDTO> ventas = await _gestionCalidadRepository.VentasPorCliente(filtros);

            if (ventas.Count < 1)
                return new ResponseModel<string>(true, "No se encontraron ventas", null);


            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);

            string reporte = VentasPorClienteReport.Exportar(logoUnilene, ventas);

            ResponseModel<string> response = new ResponseModel<string>(true, Constante.MESSSGE_SUCCESS_REPORT, reporte);
            return response;
        }


        public async Task<IEnumerable<DatosFormatoListarSsomaModel>> ListarSsoma(int TipoDocumento, string Codigo)
        {
            IEnumerable<DatosFormatoListarSsomaModel> respuesta =await _gestionCalidadRepository.ListarSsoma(TipoDocumento, Codigo);

            return respuesta;
        }

        public async Task<ResponseModel<string>> RegistrarSsoma(DatosFormatoRegistrarSsomaModel dato , string UsuarioSesion)
        {
            dynamic respuesta = new { mensaje = "", respuesta = false };

            respuesta = await _gestionCalidadRepository.RegistrarSsoma(dato, UsuarioSesion);

            return new ResponseModel<string>(respuesta.respuesta, Constante.MESSAGE_SUCCESS, respuesta.mensaje);
        }

        public async Task<ResponseModel<string>> EliminarSsoma(int idSsoma, string UsuarioSesion)
        {
            await _gestionCalidadRepository.EliminarSsoma(idSsoma, UsuarioSesion);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Eliminacion satisfactoria");
        }
        public async Task<PaginacionModel<ListaReclamosDTO>> ListarReclamosQuejas(FiltrosListaReclamosDTO filtros)
        {
            if (!filtros.ValidarFiltros() || !Shared.ValidarFecha(filtros.FechaInicio) || !Shared.ValidarFecha(filtros.FechaFin))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            (List<ListaReclamosDTO> listaReclamos, int cantidadRegistros) response = await _gestionCalidadRepository.ListarReclamosQuejas(filtros);

            return new PaginacionModel<ListaReclamosDTO>(response.listaReclamos, filtros.Pagina, filtros.RegistrosPorPagina, response.cantidadRegistros);
        }

        public async Task<ResponseModel<object>> RegistrarReclamoCabecera(int codigoCliente, string codigoUsuarioSesion)
        {
            if (codigoCliente == 0 || string.IsNullOrEmpty(codigoUsuarioSesion))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            string codigoReclamo = await ObtenerNuevoCodigoReclamo();

            TBMReclamosEntity reclamo = new TBMReclamosEntity()
            {
                CodReclamo = codigoReclamo,
                Cliente = codigoCliente,
                Estado = "P",
                UsuarioRegistro = codigoUsuarioSesion,
                FechaRegistro = DateTime.Now
            };

            await _gestionCalidadRepository.RegistrarReclamo(reclamo);

            object response = new { codigoReclamo = codigoReclamo, fechaRegistro = reclamo.FechaRegistro };

            return new ResponseModel<object>(true, Constante.MESSAGE_SUCCESS, response);
        }

        public async Task<ResponseModel<ReclamoDTO>> ObtenerDetalleReclamo(string codigoReclamo)
        {
            if (string.IsNullOrEmpty(codigoReclamo))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            ReclamoDTO detalleReclamo = new ReclamoDTO();
            detalleReclamo.Cabecera = await _gestionCalidadRepository.ObtenerCabeceraDetalleReclamo(codigoReclamo);
            detalleReclamo.Detalle = await _gestionCalidadRepository.ListarDetalleReclamo(codigoReclamo);

            return new ResponseModel<ReclamoDTO>(true, Constante.MESSAGE_SUCCESS, detalleReclamo);
        }

        public async Task<ResponseModel<IEnumerable<LotesFiltradosReclamo>>> LotesFiltradosReclamo(FiltrosLotesReclamosDTO filtros)
        {
            if (!filtros.Validacion())
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            IEnumerable<LotesFiltradosReclamo> lotes = await _gestionCalidadRepository.LotesFiltradosReclamo(filtros);

            return new ResponseModel<IEnumerable<LotesFiltradosReclamo>>(true, Constante.MESSAGE_SUCCESS, lotes);
        }

        public async Task<ResponseModel<DatosLoteReclamoDTO>> DatosItemLote(string lote)
        {
            if (string.IsNullOrEmpty(lote))
                throw new ValidationModelException("Error al obtener la configuración del código correlativo");

            DatosLoteReclamoDTO datosItem = await _gestionCalidadRepository.DatosItemLote(lote);

            return new ResponseModel<DatosLoteReclamoDTO>(true, Constante.MESSAGE_SUCCESS, datosItem);
        }

        public async Task<ResponseModel<string>> ActualizarDetalleLoteReclamo(TBDReclamosEntity detalle)
        {
            detalle.Estado = "P";

            if (!detalle.ValidarCreacion() || !Shared.ValidarFecha(detalle.FechaIncidencia))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            if ((detalle.Clasificacion ?? 0) != 0)
            {
                bool validarExisteClasificacion = await _commonRepository.ValidarExiteConfiguracionDetallePorId(detalle.Clasificacion ?? 0);

                if (!validarExisteClasificacion)
                    throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);
            }

            if ((detalle.AreaInvolucrada ?? 0) != 0)
            {
                bool validarExisteArea = await _commonRepository.ValidarExiteConfiguracionDetallePorId(detalle.AreaInvolucrada ?? 0);

                if (!validarExisteArea)
                    throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);
            }

            bool validacionExisteDetalle = await _gestionCalidadRepository.ValidarExisteDetalleReclamo(detalle.CodReclamo,
                detalle.Lote, detalle.OrdenFabricacion, detalle.TipoDocumento, detalle.Documento);

            if (!validacionExisteDetalle)
                return new ResponseModel<string>(false, "No se puedo encontrar el detalle del reclamo.", null);

            detalle.FechaRegistro = DateTime.Now;

            int registrosActualizados = await _gestionCalidadRepository.ActualizarDetalleLoteReclamo(detalle);

            if (registrosActualizados != 1)
                return new ResponseModel<string>(false, "Se han actualizado más registros de lo esperado.", null);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, null);
        }

        public async Task<ResponseModel<object>> GuardarDetalleReclamo(TBDReclamosEntity detalle)
        {
            detalle.Estado = "P";

            if (!detalle.ValidarCreacion() || !Shared.ValidarFecha(detalle.FechaIncidencia))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            if ((detalle.Clasificacion ?? 0) != 0)
            {
                bool validarExisteClasificacion = await _commonRepository.ValidarExiteConfiguracionDetallePorId(detalle.Clasificacion ?? 0);

                if (!validarExisteClasificacion)
                    throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);
            }

            if ((detalle.AreaInvolucrada ?? 0) != 0)
            {
                bool validarExisteArea = await _commonRepository.ValidarExiteConfiguracionDetallePorId(detalle.AreaInvolucrada ?? 0);

                if (!validarExisteArea)
                    throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);
            }

            bool validacionExisteDetalle = await _gestionCalidadRepository.ValidarExisteDetalleReclamo(detalle.CodReclamo,
                detalle.Lote, detalle.OrdenFabricacion, detalle.TipoDocumento, detalle.Documento);

            if (validacionExisteDetalle)
                return new ResponseModel<object>(false, "Ya existe un detalle registrado para el lote y el documento.", null);

            detalle.FechaRegistro = DateTime.Now;

            int idDetalle = await _gestionCalidadRepository.GuardarDetalleReclamo(detalle);

            return new ResponseModel<object>(true, Constante.MESSAGE_SUCCESS, new { idDetalle = idDetalle } );
        }

        public async Task<ResponseModel<string>> RegistrarRespuestaReclamo(RespuestaReclamoDTO respuesta)
        {
            if (!respuesta.ValidarRegistro())
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            string estado = await _gestionCalidadRepository.ObtenerEstadoLoteDetalleReclamo(respuesta.IdDetalle);

            if (string.IsNullOrEmpty(estado))
                throw new ValidationModelException("No se pudo encontrar el detalle del reclamo.");

            if (estado == "A" || estado == "R")
                throw new ValidationModelException("El reclamo ya esta cerrada");

            int registrosModificados = await _gestionCalidadRepository.RegistrarRespuestaReclamo(respuesta);

            if (registrosModificados == 0)
                throw new ValidationModelException("No se pudo encontrar el detalle.");

            await _gestionCalidadRepository.Validar_ActualizarEstadoReclamo(respuesta.IdDetalle);

            return new ResponseModel<string>(true, "Se registro correctamente la respuesta.", null);
        }

        private async Task<string> ObtenerNuevoCodigoReclamo()
        {
            IEnumerable<ConfiguracionEntity> configuracion = await _commonRepository.ObtenerConfiguracionesSistema(12, "RECLAMOS_QUEJAS");

            if (configuracion.Count() < 1)
                throw new ValidationModelException("Error al obtener la configuración del código correlativo");

            int codigoCorrelativo = (configuracion.FirstOrDefault().ValorEntero1 ?? 0) + 1;

            await _commonRepository.ActualizarCorrelativoCodReclamo(codigoCorrelativo, 12, 59, "RECLAMOS_QUEJAS");

            string codigoFormato = codigoCorrelativo.ToString() + "-" + DateTime.Now.Year.ToString();

            codigoFormato = codigoFormato.PadLeft(10, '0');

            return codigoFormato;
        }

        public async Task<ResponseModel<CabeceraReclamoLoteDTO>> DataReclamoLote (string codReclamo, string lote, string documento)
        {
            if(string.IsNullOrEmpty(codReclamo) || string.IsNullOrEmpty(lote) || string.IsNullOrEmpty(documento) )
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            CabeceraReclamoLoteDTO cabecera = await _gestionCalidadRepository.CabeceraReclamoLote(codReclamo, lote, documento);

            return new ResponseModel<CabeceraReclamoLoteDTO>(true, Constante.MESSAGE_SUCCESS, cabecera);
        }
    }
}