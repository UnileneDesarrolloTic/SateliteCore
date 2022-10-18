using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.GestionCalidad;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class GestionCalidadServices: IGestionCalidadServices
    {
        private readonly IGestionCalidadRepository _gestionCalidadRepository;

        public GestionCalidadServices(IGestionCalidadRepository gestionCalidadRepository)
        {
            _gestionCalidadRepository = gestionCalidadRepository;
        }

        public async Task<List<MateriaPrimaDTO>> ObtenerMateriaPrima(string tipoConsulta, string lote)
        {
            if (string.IsNullOrEmpty(tipoConsulta) || string.IsNullOrEmpty(lote))
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            if(tipoConsulta != "PT" && tipoConsulta != "MP")
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            return await _gestionCalidadRepository.ObtenerMateriaPrima(tipoConsulta, lote);
        }

        public async Task<DetalleSeguimientoLoteDTO> DetalleSeguimientoPorLote(RequestLotesDetalleDTO filtros)
        {

            if (filtros.Lotes.Count < 1 || filtros.OrdenesFabricacion.Count < 1)
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            DetalleSeguimientoLoteDTO detalle = new DetalleSeguimientoLoteDTO();

            detalle.OrdenesDeCompra = await _gestionCalidadRepository.OrdenCompraPorlote(filtros.Lotes);
            detalle.OrdenesDeFabricacion = await _gestionCalidadRepository.OrdenFabricacionPorlotes(filtros.OrdenesFabricacion);
            detalle.DocumentosPedidos = await _gestionCalidadRepository.OrdenDocumentosPedidosPorLotes(filtros.OrdenesFabricacion);
            detalle.GuiasRelacionadas = await _gestionCalidadRepository.OrdenGuiasRelacionadasPorLotes(filtros.OrdenesFabricacion);

            return detalle;
        }


        public async Task<List<VentasPorClienteDTO>> VentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            if (!filtros.ValidarDatos() || !Shared.ValidarFecha(filtros.FechaInicio) || !Shared.ValidarFecha(filtros.FechaFin) || filtros.FechaInicio > filtros.FechaFin )
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            return await _gestionCalidadRepository.VentasPorCliente(filtros);
        }

        public async Task<ResponseModel<string>> ReporteVentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            if (!filtros.ValidarDatos() || !Shared.ValidarFecha(filtros.FechaInicio) || !Shared.ValidarFecha(filtros.FechaFin) || filtros.FechaInicio > filtros.FechaFin)
                throw new ValidationModelException(Constante.MODEL_VALIDATION_FAILED);

            List<VentasPorClienteDTO> ventas = await _gestionCalidadRepository.VentasPorCliente(filtros);

            if(ventas.Count < 1)
                return new ResponseModel<string>(true, "No se encontraron ventas", null);
            

            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);

            string reporte = VentasPorClienteReport.Exportar(logoUnilene, ventas);

            ResponseModel<string> response = new ResponseModel<string>(true, Constante.MESSSGE_SUCCESS_REPORT, reporte);
            return response;
        }

    }
}
