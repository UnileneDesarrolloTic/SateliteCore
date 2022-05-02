﻿using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class ControlCalidadServices : IControlCalidadServices
    {
        private readonly IControlCalidadRepository _controlCalidadRepository;

        public ControlCalidadServices(IControlCalidadRepository controlCalidadRepository)
        {
            _controlCalidadRepository = controlCalidadRepository;
        }
        public async Task<(List<CertificadoEsterilizacionEntity>, int)> ListarCertificados(DatosListarCertificadoPaginado datos)
        {
            return await _controlCalidadRepository.ListarCertificados(datos);
        }

        public async Task<(List<LoteEntity>, int)> ListarLotes(DatosLote datos)
        {
            return await _controlCalidadRepository.ListarLotes(datos);
        }

        public bool RegistrarCertificado(CertificadoEsterilizacionEntity certificado)
        {
            return _controlCalidadRepository.RegistrarCertificado(certificado);
        }

        public async Task<int> RegistrarLote(LoteEntity lote)
        {
            return await _controlCalidadRepository.RegistrarLote(lote);
        }
        public async Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos)
        {
            return await _controlCalidadRepository.ListarCotizaciones(datos);
        }

        #region ANALISIS DE AGUJA
        public async Task<IEnumerable<AnalisisAgujaModel>> ListarAnalisisAguja(string lote)
        {
            return await _controlCalidadRepository.ListarAnalisisAguja(lote);
        }

        public async Task<IEnumerable<AnalisisAgujaModel>> ListaOrdenesCompra(string NumeroOrden)
        {
            return await _controlCalidadRepository.ListaOrdenesCompra(NumeroOrden);
        }

        public async Task<IEnumerable<AnalisisAgujaModel>> ListarCiclos(string identificador)
        {
            return await _controlCalidadRepository.ListarCiclos(identificador);
        }
        public async Task<int> RegistrarControlAgujas(ControlAgujasModel matricula)
                {
            return await _controlCalidadRepository.RegistrarControlAgujas(matricula);
        }
        #endregion


    }
}
