using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.Comercial;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class ControlCalidadServices : IControlCalidadServices
    {
        private readonly IControlCalidadRepository _controlCalidadRepository;

        public ControlCalidadServices(IControlCalidadRepository controlCalidadRepository)
        {
            _controlCalidadRepository = controlCalidadRepository;
        }
        public async Task<(List<CertificadoEsterilizacionEntity>, int)> ListarCertificados(DatosListarCertificadoPaginado datos)
        {
            return await _controlCalidadRepository.ListarCertificados(datos);
        }

        public async Task<(List<LoteEntity>, int)> ListarLotes(DatosLote datos)
        {
            return await _controlCalidadRepository.ListarLotes(datos);
        }

        public bool RegistrarCertificado(CertificadoEsterilizacionEntity certificado)
        {
            return _controlCalidadRepository.RegistrarCertificado(certificado);
        }

        public async Task<int> RegistrarLote(LoteEntity lote)
        {
            return await _controlCalidadRepository.RegistrarLote(lote);
        }
        public async Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos)
        {
            return await _controlCalidadRepository.ListarCotizaciones(datos);
        }

        public async Task<IEnumerable<FormatoEstructuraObtenerOrdenFabricacion>> ObtenerInformacionLote(string NumeroLote)
        {
            return await _controlCalidadRepository.ObtenerInformacionLote(NumeroLote);
        }

        public async Task<IEnumerable<DatosFormatoListarTransaccion>> ListarTransaccionItem(string NumeroLote, string codAlmacen)
        {
            return await _controlCalidadRepository.ListarTransaccionItem(NumeroLote, codAlmacen);
            
        }

        public async Task<ResponseModel<string>> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato,int idUsuario)
        {
            int reponse = await _controlCalidadRepository.RegistrarLoteNumeroCaja(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con éxito");
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoKardexInternoGCM>> ListarKardexInternoNumeroLote(string NumeroLote)
        {
            return await _controlCalidadRepository.ListarKardexInternoNumeroLote(NumeroLote);
        }

        public async Task<ResponseModel<string>> ActualizarKardexInternoGCM(int idKardex, string comentarios, int idUsuario)
        {
            int reponse = await _controlCalidadRepository.ActualizarKardexInternoGCM(idKardex, comentarios, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato, int idUsuario)
        {
            int reponse = await _controlCalidadRepository.RegistrarKardexInternoGCM(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> ExportarOrdenFabricacionCaja()
        {
            IEnumerable<FormatoEstructuraObtenerOrdenFabricacion> listaOrdenFabricacionCaja =  await _controlCalidadRepository.ExportarOrdenFabricacionCaja();
            ReporteOrdenFabricacionCaja ExporteOrdenFabricacionCaja = new ReporteOrdenFabricacionCaja();
            string reporte = ExporteOrdenFabricacionCaja.GenerarReporteCaja(listaOrdenFabricacionCaja);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
            return Respuesta;
        }



    }
}
