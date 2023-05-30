using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.OrdenServicio;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class OrdenServicioServices : IOrdenServicioServices
    {
        private readonly IOrdenServicioRepository _ordenServicioRepository;

        public OrdenServicioServices(IOrdenServicioRepository ordenServicioRepository)
        {
            _ordenServicioRepository = ordenServicioRepository;
        }

        public async Task<ResponseModel<IEnumerable<ListarOrdenServicioResponseDTO>>> ListarOrdenServicio(DateTime fechaInicio, DateTime fechaFin)
        {
            if (!Shared.ValidarFecha(fechaInicio) || !Shared.ValidarFecha(fechaFin))
                throw new ValidationModelException();

            IEnumerable<ListarOrdenServicioResponseDTO> ordenes = await _ordenServicioRepository.ListarOrdenServicio(fechaInicio, fechaFin);
            return new ResponseModel<IEnumerable<ListarOrdenServicioResponseDTO>>(ordenes);
        }

        public async Task<ResponseModel<IEnumerable<ListaTransportistaComboxResponse>>> ListarTransportistaCombox()
        {
            IEnumerable<ListaTransportistaComboxResponse> ordenes = await _ordenServicioRepository.ListarTransportistaCombox();
            return new ResponseModel<IEnumerable<ListaTransportistaComboxResponse>>(ordenes);
        }

        public async Task<ResponseModel<IEnumerable<DetalleOrdenServicioResponse>>> ListaDetalleOrdenServicio(int codigoOrdenServicio)
        {
            if (codigoOrdenServicio < 1)
                throw new ValidationModelException();

            IEnumerable<DetalleOrdenServicioResponse> ordenes = await _ordenServicioRepository.ListaDetalleOrdenServicio(codigoOrdenServicio);
            return new ResponseModel<IEnumerable<DetalleOrdenServicioResponse>>(ordenes);
        }

        public async Task<ResponseModel<IEnumerable<OrdenServicioGuiaRemisionResponse>>> ListaGuiaRemision(DateTime fechaInicio, DateTime fechaFin)
        {
            if (!Shared.ValidarFecha(fechaInicio) || !Shared.ValidarFecha(fechaFin))
                throw new ValidationModelException();

            IEnumerable<OrdenServicioGuiaRemisionResponse> guias = await _ordenServicioRepository.ListaGuiaRemision(fechaInicio, fechaFin);
            return new ResponseModel<IEnumerable<OrdenServicioGuiaRemisionResponse>>(guias);
        }

        public async Task<ResponseModel<string>> ModificarOrdenServicio(OrdenServicioModificadosDTO ordenes, string usuario)
        {
            if (ordenes == null || ordenes.ItemsDetalle.Count() < 1 || ordenes.idTransportista < 1)
                throw new ValidationModelException();

            int id = ordenes.ItemsDetalle.Select(x => x.Cabecera).First();

            await _ordenServicioRepository.ModificarTransportista(id, ordenes.idTransportista);

            List<OrdenServicioDetalle> existentes = ordenes.ItemsDetalle.FindAll(x => x.Id > 0);
            List<OrdenServicioDetalle> agregados = ordenes.ItemsDetalle.FindAll(x => x.Id == 0);
            List<OrdenServicioDetalle> extras = ordenes.ItemsDetalle.FindAll(x => x.Id == -1);

            if (existentes.Count > 0)
                await _ordenServicioRepository.Modificar_Peso_Bultos(existentes);

            if (agregados.Count > 0)
            {
                List<RegistrarGuia_OrdenServicioDTO> datosGuias = new List<RegistrarGuia_OrdenServicioDTO>();
                agregados.ForEach(x =>
                {
                    string serie = x.Guia.Substring(0, x.Guia.IndexOf("-") );
                    string documento = x.Guia.Substring( x.Guia.IndexOf("-") + 1 );

                    datosGuias.Add(new RegistrarGuia_OrdenServicioDTO(x.Cabecera, serie, documento, x.Peso, x.Bultos, usuario));
                });

                await _ordenServicioRepository.RegistrarGuias_OrdenServicio(datosGuias);
            }

            if (extras.Count > 0) 
            {
                List<OrdenServicioDetalle> auxiliar = extras.Select( x => new OrdenServicioDetalle(x, usuario)).ToList();
                await _ordenServicioRepository.RegistrarObjetosExtrasEnvio_OS(auxiliar);
            }


            return new ResponseModel<string>("Se modificó la orden de servicio");
        }

        public async Task<ResponseModel<string>> EliminarDetalleOrdenServicio(int id)
        {
            if (id < 1)
                throw new ValidationModelException();

            await _ordenServicioRepository.EliminarDetalleOrdenServicio(id);

            return new ResponseModel<string>("Se eliminó el detalle");
        }

        public async Task<ResponseModel<string>> EditarGuiaRemision(EditarGuiaOS_DTO datosGuia)
        {
            if (!datosGuia.Validar())
                throw new ValidationModelException();

            if(datosGuia.Id == -1)
            {
                List<OrdenServicioDetalle> datosInsertar = new List<OrdenServicioDetalle>();
                datosInsertar.Add(new OrdenServicioDetalle()
                {
                    Cabecera = datosGuia.Cabecera,
                    Id = datosGuia.Id,
                    Guia = datosGuia.Comercial,
                    Fecha = datosGuia.Fecha,
                    Cliente = datosGuia.Cliente,
                    Direccion = datosGuia.Direccion,
                    Departamento = datosGuia.Departamento,
                    Comercial = null,
                    Peso = datosGuia.Peso,
                    Bultos = datosGuia.Bultos,
                    Usuario = datosGuia.Usuario,
                });

                await _ordenServicioRepository.RegistrarObjetosExtrasEnvio_OS(datosInsertar);
            }
            else
                await _ordenServicioRepository.EditarGuiaRemision(datosGuia);
           


            return new ResponseModel<string>("Se actualizó los datos");
        }

        public async Task<ResponseModel<string>> GuardarTransportista(DatosTransportistaDTO datosTransportista)
        {
            if (string.IsNullOrWhiteSpace(datosTransportista.Descripcion))
                throw new ValidationModelException();

            await _ordenServicioRepository.GuardarTransportista(datosTransportista);

            return new ResponseModel<string>("Se ha guardado los datos del transportista");
        }

        public async Task<ResponseModel<string>> NuevaOrdenServicio(DatosRegistrarOrdenServicioDTO ordenServicio)
        {
            if (ordenServicio.Detalle.Count < 1 || ordenServicio.Transportista < 1)
                throw new ValidationModelException();

            int idCabecera = await _ordenServicioRepository.CrearOrdenServicio_Cabecera(ordenServicio.Usuario, ordenServicio.Transportista);

            List<RegistrarGuia_OrdenServicioDTO> datosGuias = new List<RegistrarGuia_OrdenServicioDTO>();
            ordenServicio.Detalle.ForEach(x =>
            {
                string serie = x.Guia.Substring(0, x.Guia.IndexOf("-"));
                string documento = x.Guia.Substring(x.Guia.IndexOf("-") + 1);

                datosGuias.Add(new RegistrarGuia_OrdenServicioDTO(idCabecera, serie, documento, x.Peso, x.Bultos, ordenServicio.Usuario));
            });

            await _ordenServicioRepository.RegistrarGuias_OrdenServicio(datosGuias);

            return new ResponseModel<string>("Se creó la Orden de Servicio...");
        }

        public async Task<ResponseModel<string>> ExportarSalidas (DateTime? inicio, DateTime? fin)
        {
            if (!Shared.ValidarFecha(inicio) || !Shared.ValidarFecha(fin))
                throw new ValidationModelException();

            List<DatosExportarSalidasDTO> datos = await _ordenServicioRepository.DatosExportarSalidas(inicio, fin);

            ReporteOrdenServicioSalidas_Excel claseReporte = new ReporteOrdenServicioSalidas_Excel(datos);
            string reporte = claseReporte.Exportar();

            if (string.IsNullOrEmpty(reporte))
                throw new Exception("Error al generar el reporte !!");

            return new ResponseModel<string>(reporte);
        }

        public async Task<ResponseModel<string>> ExportarOrdenServicio(int id)
        {
            if (id < 1)
                throw new ValidationModelException();

            DatosReporteOrdenServicioPDF_DTO datosReporte = await _ordenServicioRepository.DatosExportarOrdenServicio(id);
            ReporteOrdenServicio_PDF claseReporte = new ReporteOrdenServicio_PDF(datosReporte);
            string reporte = claseReporte.Exportar();

            if (string.IsNullOrEmpty(reporte))
                throw new Exception("Error al generar el reporte !!");

            return new ResponseModel<string>(reporte);
        }

        public async Task<ResponseModel<DatosOServicioMarcadoDTO?>> OrdenServicioRetornada(string ordenServicio)
        {
            if (string.IsNullOrWhiteSpace(ordenServicio))
                throw new ValidationModelException();

            (string os, DateTime? fechaRetorno) = await _ordenServicioRepository.ObtenerFechaSalidaOS(ordenServicio);

            if(!(fechaRetorno is null))
                return new ResponseModel<DatosOServicioMarcadoDTO?>(false, $"El retorno de la Orden de Servicio {os} fue registrada el {fechaRetorno?.ToString("dd/MM/yyyy HH:mm")}", null);

            DatosOServicioMarcadoDTO result = await _ordenServicioRepository.OrdenServicioRetornada(ordenServicio);

            if(string.IsNullOrWhiteSpace(result.OrdenServicio))
                return new ResponseModel<DatosOServicioMarcadoDTO?>(false, "No se encontro la Orden Servicio", result);

            return new ResponseModel<DatosOServicioMarcadoDTO?>(result);
        }

        public async Task<ResponseModel<string>> EliminarOrdenServicio(string ordenServicio)
        {
            if (string.IsNullOrWhiteSpace(ordenServicio))
                throw new ValidationModelException();

            await _ordenServicioRepository.EliminarOrdenServicio(ordenServicio);

            return new ResponseModel<string>("Se ha eliminado la orden de servicio");
        }

        public async Task<ResponseModel<string>> ReporteGuiaOrdenServicio(DateTime? fechaInicio, DateTime? fechaFin)
        {
            if (!Shared.ValidarFecha(fechaInicio) || !Shared.ValidarFecha(fechaFin))
                throw new ValidationModelException();

            List<DatosReporteGuiaOrdenServicioDTO> datosReporte = await _ordenServicioRepository.DatosRptGuiasOrdenServicio(fechaInicio, fechaFin);

            if(datosReporte.Count() < 1)
                return new ResponseModel<string>(false, "No se encontro registros en las fechas seleccionadas.", null);

            ReporteGuiasOrdenServicio_Excel repoteClass = new ReporteGuiasOrdenServicio_Excel(datosReporte);
            string reporte = repoteClass.Exportar();

            if (string.IsNullOrWhiteSpace(reporte))
                return new ResponseModel<string>(false, "Error al generar el reporte", null);

            return new ResponseModel<string>(reporte);
        }
    }
}
