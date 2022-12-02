using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IExportacionesRepository
    {
        public Task<IEnumerable<DatosFormatoListarCotizacionExportacion>> ListarCotizacionExportaciones(FiltrarCotizacionExportacionModel filtro);
        public Task<(object cabecera, object detalle)> BuscarCotizacionExportaciones(string NumeroDocumento);
        public Task<int> EditarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones dato, string UsuarioSesion);
        public Task<int> RegistrarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones dato, string UsuarioSesion);
        public Task<FormatoDetalleCotizacionExportaciones> ObtenerInformacionExcel(FormatoDetalleExcelExportacionesModel parametro);
        public Task<List<FormatoDetalleCotizacionExportaciones>> BuscarWHItemMast(string Opcion, string Descripcion);

        public Task<int> DesactivarItemCotizacionExportacion(string NumeroDocumento, string Item, int Linea , string UsuarioSesion);
    }
}
