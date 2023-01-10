using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.ReportServices.Contracts.Actaverifacioncc;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.ReportServices.Contracts.Dashboard;

namespace SatelliteCore.Api.Services
{
    public class LicitacionesServices : ILicitacionesServices
    {
        private readonly ILicitacionesRepository _licitacionesRepository;

        public LicitacionesServices(ILicitacionesRepository licitacionesRepository)
        {
            _licitacionesRepository = licitacionesRepository;
        }

        public async Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido)
        {
            int NumeroPedido = Int32.Parse(Pedido);

            IEnumerable<ListarDetallePedido> ListarDetalle = await _licitacionesRepository.ListaDetallePedido(NumeroPedido.ToString("D10"));
            return ListarDetalle;
        }

        public async Task<ResponseModel<string>> RegistrarProceso(DatoFormatoProcesoModel dato)
        {
            int cantidad = await _licitacionesRepository.RegistrarProceso(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "");
        }

        public async Task<IEnumerable<DatosFormatoDistribuccionLP>> ListarDistribuccionProceso(int NumeroProceso, string Item, string Mes)
        {
            return await _licitacionesRepository.ListarDistribuccionProceso(NumeroProceso, Item, Mes);
            
        }

        public async Task<IEnumerable<string>> ObtenerTipoUsuario(int NumeroProceso, int Item, string Mes)
        {
            return await _licitacionesRepository.ObtenerTipoUsuario(NumeroProceso, Item, Mes);

        }


        public async Task<ResponseModel<DatosFormatoBuscarOrdenCompraLicitacionesModel>> BuscarOrdenCompraLicitaciones(int NumeroProceso, int NumeroEntrega, int Item, string TipoUsuario)
        {
            DatosFormatoBuscarOrdenCompraLicitacionesModel response = new DatosFormatoBuscarOrdenCompraLicitacionesModel();
             response = await _licitacionesRepository.BuscarOrdenCompraLicitaciones(NumeroProceso, NumeroEntrega, Item , TipoUsuario);

            return new ResponseModel<DatosFormatoBuscarOrdenCompraLicitacionesModel>(true, Constante.MESSAGE_SUCCESS, response);

        }

        public async Task<ResponseModel<string>> RegistrarOrdenCompra(DatoFormatoRegistrarOrdenCompraLicitaciones dato, int idUsuario)
        {
            int _ = await _licitacionesRepository.RegistrarOrdenCompra(dato, idUsuario);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con exito");

        }

        public async Task<ResponseModel<string>> RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> dato)
        {
            await _licitacionesRepository.RegistrarDistribuccionProceso(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado Con Existo");
        }

        public async Task<IEnumerable<ListarProcesoEntity>> ListarProceso(int idClient)
        {
            return await _licitacionesRepository.ListarProceso(idClient);
        }

        public async Task<IEnumerable<DatosFormatoProgramacionMuestraModel>> ListarProgramaMuestraLIP(int IdProceso, string NumeroEntrega)
        {
            return await _licitacionesRepository.ListarProgramaMuestraLIP(IdProceso,NumeroEntrega);

        }
        public async Task<ResponseModel<string>> RegistrarProgramacionMuestreo(List<DatosFormatoMuestraEnsayoLIP> dato)
        {
            await _licitacionesRepository.RegistrarProgramacionMuestreo(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado Con Existo");
        }

        public async Task<ResponseModel<IEnumerable<ListarGuiaInformeLPModel>>> ListarGuiaInformacion(string NumeroEntrega, string OrdenCompra)
        {
            IEnumerable<ListarGuiaInformeLPModel> result=await _licitacionesRepository.ListarGuiaInformacion(NumeroEntrega, OrdenCompra);

            return new ResponseModel<IEnumerable<ListarGuiaInformeLPModel>>(true, Constante.MESSAGE_SUCCESS, result);  ;

        }

        public async Task<IEnumerable<EstructuraListaContratoProceso>> ListarContratoProceso(string proceso)
        {
            return await _licitacionesRepository.ListarContratoProceso(proceso);

        }

        public async Task<ResponseModel<string>> RegistrarContratoProceso(List<DatosRequestFormatoContratoProcesoModel> dato)
        {
            await _licitacionesRepository.RegistrarContratoProceso(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado Con Existo");
        }

       public async Task<ResponseModel<string>> DashboardLicitacionesExportar()
        {
            IEnumerable<DatosFormatodashboardLicitaciones> documento = await _licitacionesRepository.DashboardLicitacionesExportar();

            ReporteLicitaciones ExporteDashboard = new ReporteLicitaciones();
            string reporte = ExporteDashboard.GenerarReporteDashboardLicitaciones(documento);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }

    }
}
