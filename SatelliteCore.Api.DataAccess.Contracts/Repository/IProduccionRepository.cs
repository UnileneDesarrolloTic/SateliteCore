using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.OCDrogueria;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IProduccionRepository
    {
        public Task<SeguimientoProductoArimaModel> SeguimientoProductosArima(string periodo);
        public Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro);
        public Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla);
        public Task<IEnumerable<SeguimientoCandMPAModel>> ExportarAgujasMateriaPrima(string regla);
        public Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item);
        public Task<bool> MostrarColumnaMP(int Usuario);
        public Task<SeguimientoComprasMPArima> SeguimientoCompraMPArima(PronosticoCompraMP dato);
        public Task<FormatoEstructuraLoteEtiquetas> LoteFabricacionEtiquetas(string NumeroLote);
        public Task<int> RegistrarLoteFabricacionEtiquetas(List<DatosEstructuraLoteEtiquetasModel> dato, int idUsuario);
        public Task<IEnumerable<DatoFormatoLoteEstado>> ListarLoteEstado();
        public Task<int> ModificarLoteEstado(DatosFormatoRequestLoteEstado dato);
        public Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Origen,string Anio, string Regla);
        public Task<DatosFormatoInformacionItemOrdenCompra> BuscarItemOrdenCompra(string Item,string Anio);
        public Task<int> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra dato);
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompra(string OrdenCompra);
        public Task<int> ActualizarFechaComprometidaMasiva(DatosFormatoCabeceraOrdenCompraModel dato);
        public Task<IEnumerable<DatosFormatoReporteSeguimientoDrogueria>> SeguimientoOCDrogueria(int idproveedor);
        public Task<IEnumerable<DatosFormatoMostrarOrdenCompraDrogueria>> MostrarOrdenCompraDrogueria(string Item);
        public Task<IEnumerable<DatosFormatoMostrarProveedorDrogueria>> MostrarProveedorDrogueria();
        public Task<IEnumerable<FormatoDatosCabeceraOrdenCompraPrevio>> MostrarOrdenCompraPrevios();
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompraSimulada(string proveedor);
        public Task<int> GuardarOrdenCompraVencida(DatosFormatoCambiarEstadoOCVencida dato, string usuario);
        public Task<int> GenerarOrdenCompraDrogueria();
        public Task<int> RegistrarOrdenCompraDrogueria(DatosFormatoGuardarCabeceraOrdenCompraDrogueria dato, string strusuario, int idusuario);
        public Task<IEnumerable<DatosFormatoGestionItemDrogueriaColor>> GestionItemDrogueriaColor();
    }
}
