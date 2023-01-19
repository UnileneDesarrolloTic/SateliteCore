using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IControlCalidadRepository
    {
        public Task<(List<CertificadoEsterilizacionEntity>, int)> ListarCertificados(DatosListarCertificadoPaginado datos);
        public bool RegistrarCertificado(CertificadoEsterilizacionEntity certificado);
        public Task<(List<LoteEntity>, int)> ListarLotes(DatosLote datos);
        public Task<int> RegistrarLote(LoteEntity lote);
        public Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos);
        public Task<DatosFormatoListarOrdenFabricacionModel> ObtenerInformacionLote(string NumeroLote);
        public Task<IEnumerable<DatosFormatoListarTransaccion>> ListarTransaccionItem(string NumeroLote, string codAlmacen);
        public Task<int> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato, int idUsuario);
        public Task<IEnumerable<DatosFormatoKardexInternoGCM>> ListarKardexInternoNumeroLote(string NumeroLote);
        public Task<int> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato, int idUsuario);
        public Task<int> ActualizarKardexInternoGCM(int idKardex, string comentarios, int idUsuario);
        public Task<IEnumerable<FormatoEstructuraObtenerOrdenFabricacion>> ExportarOrdenFabricacionCaja(string anioProduccion);
        public Task<IEnumerable<DatosFormatosListarControlLotes>> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato);
        public Task<int> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato);
        public Task<IEnumerable<DatosFormatoTablaNumerodeParte>> ListarMaestroNumeroParte(string Grupo,string Tabla);
        public Task<IEnumerable<DatosFormatoTablaAbributoModel>> ListarAtributos();
        public Task<IEnumerable<DatosFormatoTablaDescripcionModel>> ListarDescripcion(string Marca, string Hebra);
        public Task<IEnumerable<DatosFormatoTablaLeyendaModel>> ListarLeyenda(string Marca, string Hebra);
        public Task<IEnumerable<DatosFormatoTablaPruebasModel>> ListarTablaPrueba(string Metodologia);
        public Task<IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel>> ListarObtenerAgujasDescripcionNuevo();
        public Task<int> NuevoDescripcionDT(DatosFormatoActualizacionDescripcionModel dato, string idUsuario);
        public Task<int> EliminarDescripcionDT(string IdDescripcion);
        public Task<IEnumerable<DatosFormatoObtenerAgujasDescripcionModel>> ListarObtenerAgujasDescripcionActualizar(string IdDescripcion);
        public Task<int> ActualizarDescripcionDT(DatosFormatoActualizacionDescripcionModel dato, string Usuario);
        public Task<int> RegistrarActualizarLeyendaDT(DatosFormatoLeyendaDTModel dato, string Usuario);
        public Task<int> EliminarLeyendaDT(string IdLeyenda);
        public Task<int> RegistrarActualizarPruebaDT(DatosFormatoNuevoPruebaModel dato, string idUsuario);
        public Task<int> EliminarPruebaDT(string IdLeyenda);
        public Task<DatosFormatoNumeroLoteProtocoloModel> BuscarNumeroLoteProtocolo(string NumeroLote,string Idioma);
        public Task<IEnumerable<DatosFormatosDatoListarPruebaProtocolo>> BuscarPruebaFormatoProtocolo(string NumeroLote,string NumeroParte, string Idioma);
        public Task<int> RegistrarControlProcesoProtocolo(DatosFormatoControlProcesosProtocoloModel dato, string Usuario);
        public Task<int> RegistrarControlPTProtocolo(DatosFormatoControlProductoTermino dato, string Usuario);
        public Task<int> RegistrarPruebasEfectuadasProtocolo(DatosFormatoPruebasEfectuasProtocolos dato, string idUsuario);
        public Task<IEnumerable<DatosFormatoInformacionResultadoProtocolo>> BuscarInformacionResultadoProtocolo(string NumeroLote);
        public Task<int> InsertarCabeceraFormatoProtocolo(DatosFormatoCabeceraFormatoProtocolo dato, string UsuarioSesion);
        public Task<IEnumerable<DatosFormatoInformacionResultadoProtocolo>> ImprimirControlProceso(string NumeroLote);
        public Task<IEnumerable<DatosFormatoProtocoloPruebaModel>> ImprimirDocumentoProtocolo(string NumeroLote, string Idioma);
        public Task<ParametroMastEntity> ProtocoloRevisionTerminado();

    }
}
