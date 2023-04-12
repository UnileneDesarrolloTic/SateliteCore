using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request;
using System.Collections.Generic;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Response.Licitaciones;
using SatelliteCore.Api.Models.Request.Licitaciones;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILicitacionesServices
    {
        public Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido);
        public Task<ResponseModel<string>> RegistrarProceso(DatoFormatoProcesoModel dato);
        public Task<IEnumerable<DatosFormatoDistribuccionLP>> ListarDistribuccionProceso(int NumeroProceso, string Item, string Mes);
        public Task<IEnumerable<string>> ObtenerTipoUsuario(int NumeroProceso, int Item, string Mes);
        public Task<ResponseModel<DatosFormatoBuscarOrdenCompraLicitacionesModel>> BuscarOrdenCompraLicitaciones(int NumeroProceso, int NumeroEntrega, int Item, string TipoUsuario);
        public Task<ResponseModel<string>> RegistrarOrdenCompra(DatoFormatoRegistrarOrdenCompraLicitaciones dato, int idUsuario);
        public Task<ResponseModel<string>> RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> dato);
        public Task<IEnumerable<ListarProcesoEntity>> ListarProceso(int idClient);
        public Task<IEnumerable<DatosFormatoProgramacionMuestraModel>> ListarProgramaMuestraLIP(int IdProceso, string NumeroEntrega);
        public Task<ResponseModel<string>> RegistrarProgramacionMuestreo(List<DatosFormatoMuestraEnsayoLIP> dato);
        public Task<ResponseModel<IEnumerable<ListarGuiaInformeLPModel>>> ListarGuiaInformacion(string NumeroEntrega, string OrdenCompra);
        public Task<IEnumerable<EstructuraListaContratoProceso>> ListarContratoProceso(string proceso);
        public Task<ResponseModel<string>> RegistrarContratoProceso(List<DatosRequestFormatoContratoProcesoModel> dato);
        public Task<ResponseModel<string>> DashboardLicitacionesExportar(string opcion);
        public Task<ResponseModel<DatosFormatoInformacionFacturaExpediente>> BuscarFacturaProceso(string factura, string usuario);
        public Task<ResponseModel<string>> RegistrarExpedienteLI(DatosFormatoRegistrarExpedienteLi dato);
    }
}
