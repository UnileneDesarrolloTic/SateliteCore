using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IDispensacionServices
    {
        public Task<ResponseModel<IEnumerable<DatosFormatoObtenerOrdenFabricacion>>> ObtenerOrdenFabricacion(DatosFormatoFiltroOrdenFabricacion dato);
        public Task<IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion>> RecetasOrdenFabricacion(string ordenFabricacion);
        public Task<ResponseModel<string>> RegistrarDispensacionMP(DatosFormatoDispensacionMateriaPrima dato, string usuario);
        public Task<IEnumerable<DatosFormatoHistorialDispensaccion>> HistorialDispensacionMP(string ordenFabricacion, string lote);
        public Task<ResponseModel<DatosFormatoInformacionDispensacionPT>> InformacionItem(string item, string ordenFabricacion, string secuencia);
        public Task<IEnumerable<DatosFormatoDispensacionRecetaDetalle>> DetalleDispensacionReceta();
        public Task<ResponseModel<string>> RegistrarRecetasGlobal(List<DatosFormatoRegistroDispensacionRecetaGlobal> dato, string usuario);
        public Task<IEnumerable<DatosFormatoDispensacionGuiaDespacho>> DispensacionGuiaDespacho();
        public Task<IEnumerable<DatosFormatoMostrarDispensacionDespacho>> MostrarDispensacionDespacho(string id);
        public ResponseModel<string> GeneracionPdfDespacho(string id);
    }
}
