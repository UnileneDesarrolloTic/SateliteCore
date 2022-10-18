using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IContabilidadRepository
    {
        public Task<IEnumerable<DetraccionesEntity>> ListarDetraccion();
        public int ProcesarDetraccionContabilidad(List<FormatoComprobantePagoDetraccion> dato);
        public Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ConsultarProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato);
        //public Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ProcesarProductoExcel(DatosFormatoFiltrarAnalisisCostoRequest dato);
        public Task<IEnumerable<DatosFormatoRecetaItemComponente>> ConsultarRecetaProducto(string Item,string FechaDocumento);
  
    }
}
