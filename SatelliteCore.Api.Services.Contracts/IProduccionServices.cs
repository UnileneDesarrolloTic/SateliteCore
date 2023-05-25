using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.OCDrogueria;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.CompraImportacion;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IProduccionServices
    {
        public Task<List<ProductoArimaModel>> SeguimientoProductosArima(string periodo);
        public Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro);
        public Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla);
        public Task<ResponseModel<string>> ExportarAgujasMateriaPrima(string regla);
        public Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item);
        public Task<bool> MostrarColumnaMP(int usuario);
        public Task<List<CompraMPArimaModel>> SeguimientoCompraMPArima(PronosticoCompraMP dato);
        public Task<ResponseModel<string>> CompraMateriaPrimaExportar(PronosticoCompraMP dato);
        public Task<ResponseModel<FormatoEstructuraLoteEtiquetas>> LoteFabricacionEtiquetas(string NumeroLote);
        public Task<ResponseModel<string>> RegistrarLoteFabricacionEtiquetas(List<DatosEstructuraLoteEtiquetasModel> dato,int idUsuario);
        public Task<IEnumerable<DatoFormatoLoteEstado>> ListarLoteEstado();
        public Task<ResponseModel<string>> ModificarLoteEstado(DatosFormatoRequestLoteEstado dato);
        public Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Anio);
        public Task<DatosFormatoInformacionItemOrdenCompra> BuscarItemOrdenCompra(string Item,string Anio);
        public Task<ResponseModel<string>> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra dato);
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompra(string OrdenCompra);
        public Task<ResponseModel<string>> ActualizarFechaPrometidaMasiva(DatosFormatoCabeceraOrdenCompraModel dato);
        public Task<ResponseModel<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>>> SeguimientoOCDrogueria(int idproveedor);
        public Task<ResponseModel<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>>> MostrarOrdenCompraDrogueria(string Item);
        public Task<ResponseModel<IEnumerable<DatosFormatoMostrarProveedorDrogueria>>> MostrarProveedorDrogueria();
        public Task<ResponseModel<string>> ExcelCompraDrogueria(int idproveedor,bool mostrarcolumna, string agrupador);
        public Task<ResponseModel<IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio>>> MostrarOrdenCompraPrevios();
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompraSimulada(string proveedor, string usuario);
        public Task<ResponseModel<string>> GuardarOrdenCompraVencida(DatosFormatoCambiarEstadoOCVencida dato, string usuario);
        public Task<DatosInformacionGeneralReporteCompraArimaAgujas> InformacionSeguimientoAguja();
        public Task<ResponseModel<string>> InformacionSeguimientoAgujaExcel(string mostrarColumna);
        public Task<ResponseModel<IEnumerable<DatosFormatoTransitoPendienteOC>>> MostrarOrdenCompraArima(string Item, string Tipo);
        public Task<ResponseModel<string>> GenerarOrdenCompraDrogueria(int idUsuario);
        public Task<ResponseModel<string>> RegistrarOrdenCompraDrogueria(DatosFormatoGuardarCabeceraOrdenCompraDrogueria dato, string strusuario, int idusuario);
        public Task<ResponseModel<IEnumerable<DatosFormatoLitadoSeguimientoCompraImportada>>> InformacionSeguimientoCompraImportacion(int material);
        public Task<ResponseModel<IEnumerable<DatosFormatoListadoCommodity>>> InformacionSeguimientoCompraCommodity();
        public Task<ResponseModel<string>> ReporteSeguimientoCompraImportacionExcel(string mostrarColumna, string reporteArima);
        public Task<ResponseModel<IEnumerable<DatosFormatoMostrarOrdenCompraNacionalImportacion>>> MostrarOrdenCompraNacionalImportacion(string item, string tipo, int material);
    }
}
