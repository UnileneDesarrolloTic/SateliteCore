﻿using Microsoft.Extensions.DependencyInjection;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.DataAccess.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Services;
using SatelliteCore.Api.Services.Contracts;

namespace SatelliteCore.Api.Config
{
    public static class IoCRegister
    {
        public static IServiceCollection AddRegistration(this IServiceCollection service)
        {
            AddRegisterRepositories(service);
            AddRegisterServices(service);
            AddRegisterOthers(service);
            return service;
        }
        private static IServiceCollection AddRegisterRepositories(this IServiceCollection service)
        {
            service.AddScoped<IUsuarioRepository, UsuarioRepository>();
            service.AddScoped<IAccesosRepository, AccesosRepository>();
            service.AddScoped<ICommonRepository, CommonRepository>();
            service.AddScoped<IProduccionRepository, ProduccionRepository>();
            service.AddScoped<IControlCalidadRepository, ControlCalidadRepository>();
            service.AddScoped<IComercialRepository, ComercialRepository>();
            service.AddScoped<ICotizacionRepository, CotizacionRepository>();
            service.AddScoped<IContabilidadRepository, ContabilidadRepository>();
            service.AddScoped<IAnalisisAgujaRepository, AnalisisAgujaRepository>();
            service.AddScoped<ILicitacionesRepository, LicitacionesRepository>();
            service.AddScoped<ILogisticaRepository, LogisticaRepository>();
            service.AddScoped<IRRHHRepository, RRHHRepository>();
            service.AddScoped<IGestionCalidadRepository, GestionCalidadRepository>();
            service.AddScoped<IExportacionesRepository, ExportacionesRepository>();
            service.AddScoped<IOrdenServicioRepository, OrdenServicioRepository>();
            service.AddScoped<IDispensacionRepository, DispensacionRepository>();
            service.AddScoped<IProgramacionOperacionesRepository, ProgramacionOperacionesRepository>();
            service.AddScoped<IAnalisisMateriaPrimaRepository, AnalisisMateriaPrimaRepository>();
            service.AddScoped<IEncajadoRespository, EncajadoRespository>();
            service.AddScoped<IRegistroAsistenciaRepository, RegistroAsistenciaRepository>();
            service.AddScoped<IGestionEquipoEngasteRepository, GestionEquipoEngasteRepository>();
            service.AddScoped<ITransferenciaPtRepository, TransferenciaPtRepository>();
            service.AddScoped<IComprobanteOrdenCompraRepository, ComprobanteOrdenCompraRepository>();

            return service;
        }
        private static IServiceCollection AddRegisterServices(this IServiceCollection service)
        {
            service.AddScoped<IUsuarioService, UsuarioService>();
            service.AddScoped<IAuthServices, AuthServices>();
            service.AddScoped<IValidacionesServices, ValidacionesServices>();
            service.AddScoped<ICommonServices, CommonServices>();
            service.AddScoped<IProduccionServices, ProduccionServices>();
            service.AddScoped<IControlCalidadServices, ControlCalidadServices>();
            service.AddScoped<IComercialServices, ComercialServices>();
            service.AddScoped<ICotizacionServices, CotizacionServices>();
            service.AddScoped<IContabilidadService, ContabilidadServices>();
            service.AddScoped<IAnalisisAgujaServices, AnalisisAgujaServices>();
            service.AddScoped<ILicitacionesServices, LicitacionesServices>();
            service.AddScoped<ILogisticaServices, LogisticaServices>();
            service.AddScoped<IRRHHServices, RRHHServices>();
            service.AddScoped<IGestionCalidadServices, GestionCalidadServices>();
            service.AddScoped<IExportacionesServices, ExportacionesServices>();
            service.AddScoped<IOrdenServicioServices, OrdenServicioServices>();
            service.AddScoped<IDispensacionServices, DispensacionServices>();
            service.AddScoped<IProgramacionOperacionesServices, ProgramacionOperacionesServices>();
            service.AddScoped<IAnalisisMateriaPrimaServices, AnalisisMateriaPrimaServices>();
            service.AddScoped<IEncajadoServices, EncajadoServices>();
            service.AddScoped<IRegistroAsistenciaServices, RegistroAsistenciaServices>();
            service.AddScoped<IGestionEquipoEngasteServices, GestionEquipoEngasteServices>();
            service.AddScoped<ITransferenciaPtServices, TransferenciaPtServices>();
            service.AddScoped<IComprobanteOrdenCompraServices, ComprobanteOrdenCompraServices>();

            return service;
        }

        private static IServiceCollection AddRegisterOthers(this IServiceCollection service)
        {
            service.AddScoped<IAppConfig, AppConfig>();
            return service;
        }
    }
}
