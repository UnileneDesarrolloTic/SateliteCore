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
        public Task<IEnumerable<FormatoEstructuraObtenerOrdenFabricacion>> ExportarOrdenFabricacionCaja();
        public Task<(List<DatosFormatosListarControlLotes>, int)> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato);
        public Task<int> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato);
    }
}
