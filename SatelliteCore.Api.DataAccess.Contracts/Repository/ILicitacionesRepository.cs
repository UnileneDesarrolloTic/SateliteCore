using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request;
using System.Collections.Generic;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Entities;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ILicitacionesRepository
    {
        public Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido);
        public Task<int> RegistrarProceso(DatoFormatoProcesoModel matricula);
        public Task<IEnumerable<DatosFormatoDistribuccionLP>> ListarDistribuccionProceso(int NumeroProceso, string Item, string Mes);
        public Task<IEnumerable<string>> ObtenerTipoUsuario(int NumeroProceso, int Item, string Mes);
        public Task<DatosFormatoBuscarOrdenCompraLicitacionesModel> BuscarOrdenCompraLicitaciones(int NumeroProceso, int NumeroEntrega, int Item, string TipoUsuario);
        public Task<int> RegistrarOrdenCompra(DatoFormatoRegistrarOrdenCompraLicitaciones dato, int idUsuario);
        public Task RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> matricula);
        public Task<IEnumerable<ListarProcesoEntity>> ListarProceso();
        public Task<IEnumerable<DatosFormatoProgramacionMuestraModel>> ListarProgramaMuestraLIP(int IdProceso, string NumeroEntrega);
        public Task RegistrarProgramacionMuestreo(List<DatosFormatoMuestraEnsayoLIP> matricula);

        public Task<IEnumerable<ListarGuiaInformeLPModel>> ListarGuiaInformacion(string NumeroEntrega,string OrdenCompra);

        public Task<IEnumerable<EstructuraListaContratoProceso>> ListarContratoProceso(string proceso);

        public Task RegistrarContratoProceso(List<DatosRequestFormatoContratoProcesoModel> matricula);

        public Task<IEnumerable<DatosFormatodashboardLicitaciones>> DashboardLicitacionesExportar();

    }
}
