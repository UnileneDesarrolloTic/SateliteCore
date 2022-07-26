using SatelliteCore.Api.Models.Entities;
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
        public Task<ResponseModel<FormatoEstructuraObtenerOrdenFabricacion>> ObtenerOrdenFabricacion(string OrdenFabricacion);
        public Task<IEnumerable<DatosFormatoListarTransaccion>> ListarTransaccionItem(string OrdenFabricacion, string codAlmacen);
        public Task<ResponseModel<string>> RegistrarOrdenFabricacionCaja(List<DatosFormatoOrdenFabricacionRequest> lote);
        public Task<ResponseModel<string>> ExportarOrdenFabricacionCaja();

    }
}
