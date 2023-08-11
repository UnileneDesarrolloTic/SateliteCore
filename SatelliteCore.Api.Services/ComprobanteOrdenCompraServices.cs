using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class ComprobanteOrdenCompraServices : IComprobanteOrdenCompraServices
    {
        private readonly IComprobanteOrdenCompraRepository _comprobanteOrdenCompraRepository;

        public ComprobanteOrdenCompraServices(IComprobanteOrdenCompraRepository comprobanteOrdenCompraRepository)
        {
            _comprobanteOrdenCompraRepository = comprobanteOrdenCompraRepository;
           
        }

        public async Task<IEnumerable<MostrarFechaPrometida>> MostrarInformacionOrdenCompra(string ordenCompra, string secuencia, string item)
        {
            return await _comprobanteOrdenCompraRepository.MostrarInformacionOrdenCompra(ordenCompra, secuencia, item);
        }

        public async Task<ResponseModel<string>> RegistrarFechaPrometida(DatosFormatoRegistrarFecha dato, string usuario)
        {
            if (string.IsNullOrWhiteSpace(dato.comentario))
                throw new ValidationModelException("verificar los parametros enviados");
            if (dato.detalle.Count == 0)
                return new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "debe elegir uno codigo o mas Item");

            await _comprobanteOrdenCompraRepository.RegistrarFechaPrometida(dato, usuario);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "registrado");
        }

        public async Task<IEnumerable<DatosFormatoDetalleOrdenCompra>> MostrarDetalleOrdenCompra(string ordenCompra, string item, string secuencia)
        {
            if (string.IsNullOrWhiteSpace(ordenCompra))
                throw new ValidationModelException("verificar los parametros enviados");

            IEnumerable<DatosFormatoDetalleOrdenCompra> listado = new List<DatosFormatoDetalleOrdenCompra>();

            listado = await _comprobanteOrdenCompraRepository.MostrarDetalleOrdenCompra(ordenCompra, item, secuencia);

            return listado;
        }

    }
}
