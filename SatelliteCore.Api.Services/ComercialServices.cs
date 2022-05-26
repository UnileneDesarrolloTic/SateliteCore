﻿using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class ComercialServices : IComercialServices
    {
        private readonly IComercialRepository _comercialRepository;

        public ComercialServices(IComercialRepository comercialRepository)
        {
            _comercialRepository = comercialRepository;
        }
        public async Task<(List<CotizacionEntity>, int)> ListarCotizaciones(DatosListarCotizacionesPaginado datos)
        {
            return await _comercialRepository.ListarCotizaciones(datos);
        }

        public async Task<FormatoCotizacionEntity> ObtenerEstructuraFormato(DatosEstructuraFormatoCotizacion datos)
        {
            return await _comercialRepository.ObtenerEstructuraFormato(datos);
        }

        public async Task<int> RegistrarRespuestas(FormatoCotizacionRespuesta datos)
        {
            return await _comercialRepository.RegistrarRespuestas(datos);
        }

        public async Task<(List<DetalleProtocoloAnalisis>, int)> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos)
        {
            return await _comercialRepository.ListarProtocoloAnalisis(datos);
        }
        public async Task<List<DetalleClientes>> ListarClientes()
        {
            return await _comercialRepository.ListarClientes();
        }

        public async Task<IEnumerable<FormatoLicitaciones>> ListarDocumentoLicitacion( DatosFormatoDocumentoLicitacion dato)
        {
            return await _comercialRepository.ListarDocumentoLicitacion(dato);
        }

        public async Task<List<CReporteGuiaRemisionModel>> NumerodeGuiaLicitacion(List<FormatoLicitacionesOT> dato)
        {
            StringBuilder builder = new StringBuilder();
            foreach (var safePrime in dato)
            {
                builder.Append("'"+safePrime.GuiasNumero+"'").Append(",");
            }
            string result = builder.ToString();
            string final = result.Remove(result.Length - 1);

            FormatoReporteGuiaRemisionesModel respuesta = await _comercialRepository.NumerodeGuiaLicitacion(final);
            List<DReportGuiaRemisionModel> aux = null;

           foreach(CReporteGuiaRemisionModel guia in respuesta.CabeceraReporteGuiaRemision)
            {
                aux = null;
                aux = respuesta.DetalleReporteGuiaRemision.FindAll(x => x.Guia == guia.GuiaNumero);

                if (aux.Count > 0)
                    guia.DetalleGuia.AddRange(aux);
            }

         
                    
         
            return respuesta.CabeceraReporteGuiaRemision;
        }


    }
}
