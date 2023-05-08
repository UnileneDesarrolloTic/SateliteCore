using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.OCDrogueria;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using SatelliteCore.Api.ReportServices.Contracts.Produccion;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class ProduccionServices : IProduccionServices
    {
        private readonly IProduccionRepository _pronosticoRepository;
        private readonly ICommonRepository _commonRepository;

        public ProduccionServices(IProduccionRepository pronosticoRepository, ICommonRepository commonRepository)
        {
            _pronosticoRepository = pronosticoRepository;
            _commonRepository = commonRepository;
        }

        public async Task<List<ProductoArimaModel>> SeguimientoProductosArima(string periodo)
        {
            SeguimientoProductoArimaModel productosArima = await _pronosticoRepository.SeguimientoProductosArima(periodo);
            List<TransitoProductoArimaModel> aux = null;            

            foreach (ProductoArimaModel pronostico in productosArima.Productos)
            {
                aux = null;
                aux = productosArima.DetalleTransito.FindAll(x => x.Item == pronostico.Item);

                if (aux.Count > 0)
                    pronostico.PedidosTransito.AddRange(aux);
            }

            return productosArima.Productos;
        }

        public async Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro)
        {
            return await _pronosticoRepository.ListaPedidosCreadoAuto(filtro);
        }

        public async Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla)
        {
            return await _pronosticoRepository.ListaSeguimientoCandidatosMP(regla);
        }

        public async Task<ResponseModel<string>> ExportarAgujasMateriaPrima(string regla)
        {
            IEnumerable<SeguimientoCandMPAModel> listar = new List<SeguimientoCandMPAModel>();
            listar= await _pronosticoRepository.ExportarAgujasMateriaPrima(regla);

            ReportExcelMateriaPrima ExcelMateriaP = new ReportExcelMateriaPrima();
            string excel = ExcelMateriaP.ReporteExcel(listar);

            return  new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, excel);
        }

        public async Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item)
        {
            List<DetalleControlCalidadItemMP> productosArima = await _pronosticoRepository.ControlCalidadItemMP(Item);
      
            decimal acumulador = 0;
            string NumeroOrden = "";
            int cnt=1;
            List<DetalleControlCalidadItemMP> aux = new List<DetalleControlCalidadItemMP>(); // List<Author> authors = new List<Author>  
            List<DetalleControlCalidadItemMP> aux2 = new List<DetalleControlCalidadItemMP>(); // List<Author> authors = new List<Author>  
            foreach (DetalleControlCalidadItemMP detalle in productosArima)
            {
                if (cnt == 1)
                {
                    NumeroOrden = detalle.ReferenciaNumeroDocumentoOrden;
                    acumulador = detalle.Cantidad;
                }
                else
                {
                    acumulador = detalle.Cantidad + acumulador;
                    if (acumulador == 0)
                    {
                        aux = new List<DetalleControlCalidadItemMP>();
                    }
                    else
                    {
                            aux.Add(detalle);
                    }
                }
                    cnt++;

            }

                 
                return aux;
        }

        public async Task<bool> MostrarColumnaMP(int usuario)
        {
            return await _pronosticoRepository.MostrarColumnaMP(usuario);
        }

        public async Task<List<CompraMPArimaModel>> SeguimientoCompraMPArima(PronosticoCompraMP dato)
        {



            SeguimientoComprasMPArima productosMPArima = await _pronosticoRepository.SeguimientoCompraMPArima(dato);
            List<DCompraMPArimaModel> aux = null;
            List<CompraMPArimaDetalleControlCalidad> auxcalidad = null;


            foreach (CompraMPArimaModel pronostico in productosMPArima.Productos)
            {
                aux = null;
                auxcalidad = null;
                aux = productosMPArima.DetalleTransito.FindAll(x => x.Item == pronostico.Item);
                auxcalidad = productosMPArima.DetalleCalidad.FindAll(x => x.Item == pronostico.Item);

                if (aux.Count > 0)
                    pronostico.DetalleCompra.AddRange(aux);

                if (auxcalidad.Count > 0)
                    pronostico.DetalleCalidad.AddRange(auxcalidad);
            }

            return productosMPArima.Productos;
        }

        public async Task<ResponseModel<string>> CompraMateriaPrimaExportar(PronosticoCompraMP dato)
        {
            SeguimientoComprasMPArima productosMPArima = await _pronosticoRepository.SeguimientoCompraMPArima(dato);
            List<DCompraMPArimaModel> aux = null;
            List<CompraMPArimaDetalleControlCalidad> auxcalidad = null;
            foreach (CompraMPArimaModel pronostico in productosMPArima.Productos)
            {
                aux = null;
                auxcalidad = null;
                aux = productosMPArima.DetalleTransito.FindAll(x => x.Item == pronostico.Item);
                auxcalidad = productosMPArima.DetalleCalidad.FindAll(x => x.Item == pronostico.Item);

                if (aux.Count > 0)
                    pronostico.DetalleCompra.AddRange(aux);

                if (auxcalidad.Count > 0)
                    pronostico.DetalleCalidad.AddRange(auxcalidad);
            }

            ReporteExcelCompraArima ExporteCompraArima = new ReporteExcelCompraArima();
            string reporte = ExporteCompraArima.GenerarReporte(productosMPArima);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }


        public async Task<ResponseModel<FormatoEstructuraLoteEtiquetas>> LoteFabricacionEtiquetas(string NumeroLote)
        {
            FormatoEstructuraLoteEtiquetas response = await _pronosticoRepository.LoteFabricacionEtiquetas(NumeroLote);
            return new ResponseModel<FormatoEstructuraLoteEtiquetas>(true, Constante.MESSAGE_SUCCESS, response);
        }


        public async Task<ResponseModel<string>> RegistrarLoteFabricacionEtiquetas(List<DatosEstructuraLoteEtiquetasModel> dato,int idUsuario)
        {
            int response = await _pronosticoRepository.RegistrarLoteFabricacionEtiquetas(dato, idUsuario);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con exito");
        }

        public async Task<IEnumerable<DatoFormatoLoteEstado>> ListarLoteEstado()
        {
            return await _pronosticoRepository.ListarLoteEstado();
        }

        public async Task<ResponseModel<string>> ModificarLoteEstado(DatosFormatoRequestLoteEstado dato)
        {
            int response = await _pronosticoRepository.ModificarLoteEstado(dato);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Modificación con exito");
        }

        public async Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Anio)
        {
            return await _pronosticoRepository.ListarItemOrdenCompra(Anio);
        }

        public async Task<DatosFormatoInformacionItemOrdenCompra> BuscarItemOrdenCompra(string Item,string Anio)
        {
            return await _pronosticoRepository.BuscarItemOrdenCompra(Item,Anio);
        }

        public async Task<ResponseModel<string>> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra dato)
        {
                   await _pronosticoRepository.ActualizarFechaPrometida(dato);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Modificación con exito");
        }

        public async Task<(object cabecera, object detalle)> VisualizarOrdenCompra(string OrdenCompra)
        {
            (object cabecera, object detalle)  response = await _pronosticoRepository.VisualizarOrdenCompra(OrdenCompra);
            return response;
        }

        public async Task<ResponseModel<string>> ActualizarFechaPrometidaMasiva(DatosFormatoCabeceraOrdenCompraModel dato)
        {
            await _pronosticoRepository.ActualizarFechaComprometidaMasiva(dato);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Modificación con exito");
        }

        public async Task<ResponseModel<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>>> SeguimientoOCDrogueria(int idproveedor)
        {
            IEnumerable<DatosFormatoReporteSeguimientoDrogueria> resultado = new List<DatosFormatoReporteSeguimientoDrogueria>();
            resultado= await _pronosticoRepository.SeguimientoOCDrogueria(idproveedor);
            if (resultado.Count()==0)
                return new ResponseModel<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>>(false, "No hay Item a comprar", resultado);
            return new ResponseModel<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>>(true,Constante.MESSAGE_SUCCESS,resultado);
        }

        public async Task<ResponseModel<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>>> MostrarOrdenCompraDrogueria(string Item)
        {
            IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria> resultado = new List<DatosFormatoMostrarOrdenCompraDrogueria>();
            resultado = await _pronosticoRepository.MostrarOrdenCompraDrogueria(Item);
            return new ResponseModel<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>>(true, Constante.MESSAGE_SUCCESS, resultado);
        }

        public async Task<ResponseModel<IEnumerable<DatosFormatoMostrarProveedorDrogueria>>> MostrarProveedorDrogueria()
        {
            IEnumerable<DatosFormatoMostrarProveedorDrogueria> resultado = new List<DatosFormatoMostrarProveedorDrogueria>();
            resultado = await _pronosticoRepository.MostrarProveedorDrogueria();
            return new ResponseModel<IEnumerable<DatosFormatoMostrarProveedorDrogueria>>(true, Constante.MESSAGE_SUCCESS, resultado);
        }

        public async Task<ResponseModel<string>> ExcelCompraDrogueria(int idproveedor, bool mostrarcolumna, string agrupador)
        {
            IEnumerable<DatosFormatoReporteSeguimientoDrogueria> result = new List<DatosFormatoReporteSeguimientoDrogueria>();
            IEnumerable<DatosFormatoGestionItemDrogueriaColor> condicionesgestion = new List<DatosFormatoGestionItemDrogueriaColor>();

            result = await _pronosticoRepository.SeguimientoOCDrogueria(idproveedor);
            if (result.Count() == 0)
                return new ResponseModel<string>(false, "No hay informacion para exportar a excel", "");


            condicionesgestion = await _pronosticoRepository.GestionItemDrogueriaColor();
            ReporteExcelCompraDrogueria ExporteCompraDrogueria = new ReporteExcelCompraDrogueria();
            string reporte = ExporteCompraDrogueria.GenerarReporte(result, mostrarcolumna, condicionesgestion, agrupador);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
        }

        public async Task<ResponseModel<IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio>>> MostrarOrdenCompraPrevios()
        {
            IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio> resultado = new List<FormatoDatosCabeceraOrdenCompraPrevio>();
            resultado = await _pronosticoRepository.MostrarOrdenCompraPrevios();
            return new ResponseModel<IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio>>(true, Constante.MESSAGE_SUCCESS, resultado);
        }

        public async Task<(object cabecera, object detalle)> VisualizarOrdenCompraSimulada(string proveedor, string usuario)
        {
            if (string.IsNullOrEmpty(proveedor))
                throw new ValidationModelException("El proveedor es obligatorio");

            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.CD_INGRESAR_FORMULARIO_COMPRA_DROGUERIA;
            evento.Usuario = usuario;
            evento.Opcional = "idProveedor: " + proveedor;
            await _commonRepository.RegistroLogEvento(evento);

            (object cabecera, object detalle) response = await _pronosticoRepository.VisualizarOrdenCompraSimulada(proveedor);
            return response;
        }

        public async Task<ResponseModel<string>> GuardarOrdenCompraVencida(DatosFormatoCambiarEstadoOCVencida dato, string usuario)
        {
            await _pronosticoRepository.GuardarOrdenCompraVencida(dato, usuario);
            return new ResponseModel<string>(true, "La " + dato.numeroOrden + " con el Item " + dato.item + " ha sido excluida tránsito", "");
        }

        public async Task<ResponseModel<string>> GenerarOrdenCompraDrogueria(int idUsuario)
        {
            bool permitir=true;

            permitir = await _commonRepository.ValidacionPermisoAccesso("FR0013", idUsuario);

            if (!permitir)
            {
                await _pronosticoRepository.GenerarOrdenCompraDrogueria();
                return new ResponseModel<string>(true, "Actualizo las ordenes de compra", "");
            }
                return new ResponseModel<string>(false, "No cuenta con permiso para actualizar ordenes de compra", "");
        }
           
          

        public async Task<ResponseModel<string>> RegistrarOrdenCompraDrogueria(DatosFormatoGuardarCabeceraOrdenCompraDrogueria dato, string strusuario, int idusuario)
        {
            string ordencompra = "";
            if (dato.detalle.Count == 0)
                return new ResponseModel<string>(false, "Debe conter aunque sea un item en la orden de compra", "");

            ordencompra = await _pronosticoRepository.RegistrarOrdenCompraDrogueria(dato, strusuario, idusuario);

            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.CD_GUARDAR_ORDEN_COMPRA_DROGUERIA;
            evento.Usuario = strusuario;
            evento.Opcional = "CantidadItems: " + dato.detalle.Count.ToString() + "| OrdenCompra: " + ordencompra;
            await _commonRepository.RegistroLogEvento(evento);

            
            return new ResponseModel<string>(true, "Se registro la orden de compra", "");
        }

    }
}
