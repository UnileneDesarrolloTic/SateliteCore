﻿using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Services
{
    public class LicitacionesServices : ILicitacionesServices
    {
        private readonly ILicitacionesRepository _licitacionesRepository;

        public LicitacionesServices(ILicitacionesRepository licitacionesRepository)
        {
            _licitacionesRepository = licitacionesRepository;
        }

        public async Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido, int idCliente)
        {
            int NumeroPedido = Int32.Parse(Pedido);

            IEnumerable<ListarDetallePedido> ListarDetalle = await _licitacionesRepository.ListaDetallePedido(NumeroPedido.ToString("D10"), idCliente);
            return ListarDetalle;
        }

    }
}
