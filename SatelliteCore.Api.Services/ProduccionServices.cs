using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using SatelliteCore.Api.ReportServices.Contracts.Produccion;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class ProduccionServices : IProduccionServices
    {
        private readonly IProduccionRepository _pronosticoRepository;

        public ProduccionServices(IProduccionRepository pronosticoRepository)
        {
            _pronosticoRepository = pronosticoRepository;
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
            string excel = ExcelMateriaP.ReporteExcel(listar,regla);

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

        public async Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Origen, string Anio, string Regla)
        {
            return await _pronosticoRepository.ListarItemOrdenCompra(Origen, Anio, Regla);
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

        public async Task<ResponseModel<IEnumerable<FormatoDatosProveedorDrogueria>>> MostrarProveedorDrogueria (string id)
        {
            IEnumerable<FormatoDatosProveedorDrogueria> resultado = new List<FormatoDatosProveedorDrogueria>();
            resultado= await _pronosticoRepository.MostrarProveedorDrogueria(id);
            return new ResponseModel<IEnumerable<FormatoDatosProveedorDrogueria>>(true,Constante.MESSAGE_SUCCESS,resultado);
        }

    }
}
