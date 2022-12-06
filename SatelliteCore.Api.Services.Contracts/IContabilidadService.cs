using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Services.Contracts
{
    public interface IContabilidadService
    {
        public Task<IEnumerable<DetraccionesEntity>> ListarDetraccion();
        public int ProcesarDetraccionContabilidad(DatosFormato64 dato);
        public string GenerarBlogNotasDetraccion(FormatoProcesoDetracciones dato);
        public Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ConsultarProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato);
        public Task<ResponseModel<string>> ExportarExcelProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato);
        public Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ProcesarProductoExcel(DatosFormatoFiltrarAnalisisCostoRequest dato);
        public Task<IEnumerable<DatosFormatoRecetaItemComponente>> ConsultarRecetaProducto(string Item);
        public Task<IEnumerable<DatosFormatoComponentePrecioUnitario>> ListarItemComponentePrecio(DatosFormatosComponentPrecio dato);

    }
}
