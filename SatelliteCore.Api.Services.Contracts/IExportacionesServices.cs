using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IExportacionesServices
    {
        public Task<IEnumerable<DatosFormatoListarCotizacionExportacion>> ListarCotizacionExportaciones(FiltrarCotizacionExportacionModel filtro);
        public Task<(object cabecera, object detalle)> BuscarCotizacionExportaciones(string NumeroDocumento);
        public Task<ResponseModel<string>> GuardarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones datos, string UsuarioSesion);
        public Task<ResponseModel<List<FormatoDetalleCotizacionExportaciones>>> ProcesarExcelExportaciones(DatosFormato64 dato);
        public Task<ResponseModel<List<FormatoDetalleCotizacionExportaciones>>> BuscarWHItemMast(string Opcion, string Descripcion);
        public Task<ResponseModel<string>> DesactivarItemCotizacionExportacion(string NumeroDocumento, string Item, int Linea, string UsuarioSesion);
    }
}
