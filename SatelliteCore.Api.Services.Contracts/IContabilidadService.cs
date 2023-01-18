using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Contabildad;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Contabilidad;
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
        public Task<InformacionTransaccionKardex> InformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato);
        public Task<ResponseModel<string>> RegistrarInformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato, string usuario);
        public Task<IEnumerable<FormatoDatosCierreHistorico>> ListarInformacionReporteCierre(string Periodo);
        public Task<IEnumerable<FormatoListadoInformacionTransaccionKardex>> ListarDetalleReporteCierre(int Id, string Periodo, string Tipo);
        public Task<ResponseModel<string>> AnularReporteCierre(int Id, string usuario);

    }
}
