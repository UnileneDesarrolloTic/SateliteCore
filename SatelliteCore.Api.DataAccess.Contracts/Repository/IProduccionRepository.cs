using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
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
    }
}
