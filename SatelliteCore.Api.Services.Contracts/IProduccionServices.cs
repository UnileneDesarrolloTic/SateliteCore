using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.OCDrogueria;
using SatelliteCore.Api.Models.Response;
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
        public Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Origen, string Anio, string Regla);
        public Task<DatosFormatoInformacionItemOrdenCompra> BuscarItemOrdenCompra(string Item,string Anio);
        public Task<ResponseModel<string>> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra dato);
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompra(string OrdenCompra);
        public Task<ResponseModel<string>> ActualizarFechaPrometidaMasiva(DatosFormatoCabeceraOrdenCompraModel dato);
        public Task<ResponseModel<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>>> SeguimientoOCDrogueria(int idproveedor);
        public Task<ResponseModel<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>>> MostrarOrdenCompraDrogueria(string Item);
        public Task<ResponseModel<IEnumerable<DatosFormatoMostrarProveedorDrogueria>>> MostrarProveedorDrogueria();
        public Task<ResponseModel<string>> ExcelCompraDrogueria(int idproveedor,bool mostrarcolumna);
        public Task<ResponseModel<IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio>>> MostrarOrdenCompraPrevios();
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompraSimulada(string proveedor);
        public Task<ResponseModel<string>> GuardarOrdenCompraVencida(DatosFormatoCambiarEstadoOCVencida dato, string usuario);
    }
}
