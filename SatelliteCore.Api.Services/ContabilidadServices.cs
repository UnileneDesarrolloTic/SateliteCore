
using OfficeOpenXml;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Services
{
    public class ContabilidadServices : IContabilidadService
    {
        private readonly IContabilidadRepository _contabilidadRepository;

        public ContabilidadServices(IContabilidadRepository contabilidadRepository)
        {
            _contabilidadRepository = contabilidadRepository;
        }
        public async Task<List<DetraccionesEntity>> ListarDetraccion()
        {

            List<DetraccionesEntity> lista  = await _contabilidadRepository.ListarDetraccion();
            return lista;
        }

        public async Task<int> ProcesarDetraccionContabilidad (List<FormatoComprobantePagoDetraccion> dato)
        {
            int response = await _contabilidadRepository.ProcesarDetraccionContabilidad(dato);
            return response;
        }

      



        }
    }
