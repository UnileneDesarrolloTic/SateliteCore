using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IControlCalidadServices
    {
        public Task<(List<CertificadoEsterilizacionEntity>, int)> ListarCertificados(DatosListarCertificadoPaginado datos);
        public bool RegistrarCertificado(CertificadoEsterilizacionEntity certificado);
        public Task<(List<LoteEntity>, int)> ListarLotes(DatosLote datos);
        public Task<int> RegistrarLote(LoteEntity lote);
        public Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos);
        public Task<DatosFormatoListarOrdenFabricacionModel> ObtenerInformacionLote(string NumeroLote);
        public Task<IEnumerable<DatosFormatoListarTransaccion>> ListarTransaccionItem(string NumeroLote, string codAlmacen);
        public Task<ResponseModel<string>> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato,int idUsuario);
        public Task<IEnumerable<DatosFormatoKardexInternoGCM>> ListarKardexInternoNumeroLote(string NumeroLote);
        public Task<ResponseModel<string>> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato, int idUsuario);
        public Task<ResponseModel<string>> ActualizarKardexInternoGCM(int idKardex, string comentarios, int idUsuario);
        public Task<ResponseModel<string>> ExportarOrdenFabricacionCaja(string anioProduccion);
        public Task<IEnumerable<DatosFormatosListarControlLotes>> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato);
        public Task<ResponseModel<string>> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato);
        public Task<IEnumerable<DatosFormatoTablaNumerodeParte>> ListarMaestroNumeroParte(string Grupo, string Tabla,string Usuario);
        public Task<IEnumerable<DatosFormatoTablaAbributoModel>> ListarAtributos();
        public Task<IEnumerable<DatosFormatoTablaDescripcionModel>> ListarDescripcion(string Marca,string Hebra,string Usuario);
        public Task<IEnumerable<DatosFormatoTablaLeyendaModel>> ListarLeyenda(string Marca, string Hebra,string Usuario);
        public Task<IEnumerable<DatosFormatoTablaPruebasModel>> ListarTablaPrueba(string Metodologia,string Usuario);
        public Task<IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel>> ListarObtenerAgujasDescripcionNuevo();
        public Task<ResponseModel<string>> NuevoDescripcionDT(DatosFormatoActualizacionDescripcionModel dato,string idUsuario);
        public Task<ResponseModel<string>> EliminarDescripcionDT(string IdDescripcion,string Usuario);
        public Task<IEnumerable<DatosFormatoObtenerAgujasDescripcionModel>> ListarObtenerAgujasDescripcionActualizar(string IdDescripcion);
        public Task<ResponseModel<string>> ActualizarDescripcionDT(DatosFormatoActualizacionDescripcionModel dato, string Usuario);
        public Task<ResponseModel<string>> RegistrarActualizarLeyendaDT(DatosFormatoLeyendaDTModel dato, string Usuario);
        public Task<ResponseModel<string>> EliminarLeyendaDT(string IdLeyenda,string Usuario);
        public Task<ResponseModel<string>> RegistrarActualizarPruebaDT(DatosFormatoNuevoPruebaModel dato, string idUsuario);
        public Task<ResponseModel<string>> EliminarPruebaDT(string IdPrueba);
        public Task<ResponseModel<DatosFormatoNumeroLoteProtocoloModel>> BuscarNumeroLoteProtocolo(string NumeroLote, string Idioma);
        public Task<IEnumerable<DatosFormatosDatoListarPruebaProtocolo>> BuscarPruebaFormatoProtocolo(string NumeroLote, string NumeroParte, string Idioma);
        public Task<ResponseModel<string>> RegistrarControlProcesoProtocolo(DatosFormatoControlProcesosProtocoloModel dato,string Usuario);
        public Task<ResponseModel<string>> RegistrarControlPTProtocolo(DatosFormatoControlProductoTermino dato, string Usuario);
        public Task<ResponseModel<string>> RegistrarPruebasEfectuadasProtocolo(DatosFormatoPruebasEfectuasProtocolos dato, string idUsuario);
        public Task<IEnumerable<DatosFormatoInformacionResultadoProtocolo>> BuscarInformacionResultadoProtocolo(string NumeroLote);
        public Task<ResponseModel<string>> InsertarCabeceraFormatoProtocolo(DatosFormatoCabeceraFormatoProtocolo dato, string UsuarioSesion);
        public Task<ResponseModel<string>> ImprimirControlProcesoInterno(string NumeroLote,string UsuarioSesion);
        public Task<ResponseModel<string>> ImprimirControlPruebas(string NumeroLote,string UsuarioSesion);
        public Task<ResponseModel<string>> ImprimirDocumentoProtocolo(string NumeroLote, bool Opcion, string Idioma,string UsuarioSesion);
      
    }
}
