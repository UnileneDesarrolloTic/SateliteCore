using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.Comercial;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using SatelliteCore.Api.Models.Generic;

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

        public async Task<DatosFormatoListarOrdenFabricacionModel> ObtenerInformacionLote(string NumeroLote)
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

        public async Task<ResponseModel<string>> ExportarOrdenFabricacionCaja(string anioProduccion)
        {
            IEnumerable<FormatoEstructuraObtenerOrdenFabricacion> listaOrdenFabricacionCaja =  await _controlCalidadRepository.ExportarOrdenFabricacionCaja(anioProduccion);
            ReporteOrdenFabricacionCaja ExporteOrdenFabricacionCaja = new ReporteOrdenFabricacionCaja();
            string reporte = ExporteOrdenFabricacionCaja.GenerarReporteCaja(listaOrdenFabricacionCaja);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
            return Respuesta;
        }


        public async Task<IEnumerable<DatosFormatosListarControlLotes>> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato)
        {
            IEnumerable<DatosFormatosListarControlLotes> lista = await _controlCalidadRepository.ListarControlLotes(dato);
            return lista;
        }

        public async Task<ResponseModel<string>> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato)
        {
            int reponse = await _controlCalidadRepository.ActualizarControlLotes(dato);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoTablaNumerodeParte>> ListarMaestroNumeroParte(string Grupo, string Tabla)
        {
            IEnumerable<DatosFormatoTablaNumerodeParte> result = new List<DatosFormatoTablaNumerodeParte>();

            IEnumerable <DatosFormatoTablaNumerodeParte> Respuesta = await _controlCalidadRepository.ListarMaestroNumeroParte(Grupo,Tabla);

            foreach (DatosFormatoTablaNumerodeParte valor in Respuesta)
            {
                DatosFormatoTablaNumerodeParte myObj = new DatosFormatoTablaNumerodeParte();
                myObj.Grupo = valor.Grupo.Trim();
                myObj.NombreGrupo = valor.NombreGrupo.Trim();
                myObj.CodigoTabla = valor.CodigoTabla.Trim();
                myObj.NombreTabla = valor.NombreTabla.Trim();
                myObj.Codigo = valor.Codigo.Trim();
                myObj.DescripcionLocal = valor.DescripcionLocal.Trim();
                myObj.Longitud = valor.Longitud;
                myObj.Estado = valor.Estado;
                myObj.UltimaFechaModif = valor.UltimaFechaModif;
                myObj.UltimoUsuario = valor.UltimoUsuario;
                result=result.Concat(new[] { myObj });
            }
            
            return result;
        }


    }
}
