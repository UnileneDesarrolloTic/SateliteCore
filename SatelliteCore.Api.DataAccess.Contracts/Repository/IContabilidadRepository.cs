using MongoDB.Bson;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Contabildad;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Contabilidad;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IContabilidadRepository
    {
        public Task<IEnumerable<DetraccionesEntity>> ListarDetraccion();
        public int ProcesarDetraccionContabilidad(List<FormatoComprobantePagoDetraccion> dato);
        public Task<IEnumerable<DatosFormatoDatosProductoCostobase>> ConsultarProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato);
        public Task<IEnumerable<DatosFormatoRecetaItemComponente>> ConsultarRecetaProducto(string Item);
        public Task<IEnumerable<DatosFormatoComponentePrecioUnitario>> ListarItemComponentePrecio(DatosFormatosComponentPrecio dato);
        public Task<(List<FormatoListadoInformacionTransaccionKardex>, int)> InformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato);
        public Task<string> RegistrarInformacionTransaccionKardex(DatoFormatoRegistrarTransaccionKardex docRegistrado);
        public Task GuardarInformacionTransaccionKardex(string idMongoDB,string Tipo, string Periodo,bool CheckCierre, string usuario);
    }
}
