using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class AnalisisAgujaServices : IAnalisisAgujaServices
    {
        private readonly IAnalisisAgujaRepository _analisisAgujaRepository;

        public AnalisisAgujaServices(IAnalisisAgujaRepository analisisAgujaRepository)
        {
            _analisisAgujaRepository = analisisAgujaRepository;
        }

        public async Task<IEnumerable<ListarAnalisisAgujaModel>> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina)
        {
            ordenCompra = string.IsNullOrEmpty(ordenCompra) ? null : ("%" + ordenCompra + "%");
            lote = string.IsNullOrEmpty(lote) ? null : ("%" + lote + "%");
            item = string.IsNullOrEmpty(item) ? null : ("%" + item + "%");

            return await _analisisAgujaRepository.ListarAnalisisAguja(ordenCompra, lote, item, pagina);
        }

        public async Task<IEnumerable<ListarOrdenCompra>> ListaOrdenesCompra(string numeroOrden)
        {
            if (string.IsNullOrEmpty(numeroOrden))
                numeroOrden = "";

            numeroOrden = "000000" + numeroOrden;
            numeroOrden = "FOR" + numeroOrden.Substring((numeroOrden.Length - 6), 6);

            return await _analisisAgujaRepository.ListaOrdenesCompra(numeroOrden);
        }

        public async Task<ResponseModel<int>> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia)
        {
            int cantidad = await _analisisAgujaRepository.CantidadPruebaFlexionPorItem(controlNumero, secuencia);
            return new ResponseModel<int>(true, Constante.MESSAGE_SUCCESS, cantidad);
        }

        //public async Task<IEnumerable<ListarAnalisisAgujaModel>> ListarCiclos(string identificador)
        //{
        //    return await _analisisAgujaRepository.ListarCiclos(identificador);
        //}

        //public async Task<int> RegistrarControlAgujas(ControlAgujasModel matricula)
        //{
        //    return await _analisisAgujaRepository.RegistrarControlAgujas(matricula);
        //}
    }
}
