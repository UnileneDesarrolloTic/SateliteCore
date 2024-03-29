﻿using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.AsignacionPersonal;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.AsignacionPersonal;
using SatelliteCore.Api.Models.Response.RRHH.AsignacionPersonal;
using SatelliteCore.Api.ReportServices.Contracts.Administracion;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class UsuarioService : IUsuarioService
    {
        private readonly IUsuarioRepository _usuarioRepository;

        public UsuarioService(IUsuarioRepository usuarioRepository)
        {
            _usuarioRepository = usuarioRepository;
        }

        public async Task<UsuarioEntity> ObtenerUsuario(DatoUsuarioBasico datos)
        {
            return await _usuarioRepository.ObtenerUsuario(datos);
        }

        public async Task<AuthResponse> ObtenerUsuarioLogin(AuthRequestModel datosUsuario)
        {
            datosUsuario.Clave = Encryptations.EncryptSha256(datosUsuario.Clave);
            return await _usuarioRepository.ObtenerUsuarioLogin(datosUsuario);
        }

        public async Task<int> CambiarClave(ActualizarClaveModel datos)
        {
            datos.Clave = Encryptations.EncryptSha256(datos.Clave);
            return await _usuarioRepository.CambiarClave(datos);
        }

        public async Task<PaginacionModel<UsuarioEntity>> ListarUsuarios(DatosListarUsuarioPaginado datos)
        {

            (List<UsuarioEntity> lista, int totalRegistros) resultDb = await _usuarioRepository.ListarUsuarios(datos);

            PaginacionModel<UsuarioEntity> response = new PaginacionModel<UsuarioEntity>(resultDb.lista, datos.Pagina, datos.RegistrosPorPagina, resultDb.totalRegistros);

            return response;
        }


        public async Task<DatosFormatoAsignacionPersonalLaboralModel> ListarAsignacionPersonal()
        {
            DatosFormatoAsignacionPersonalLaboralModel response = await _usuarioRepository.ListarAsignacionPersonal();
            return response;
        }

        public async Task<List<AreaPersonalLaboralEntity>> ListarAreaPersonaLaboral()
        {
            List<AreaPersonalLaboralEntity> response = await _usuarioRepository.ListarAreaPersonaLaboral();
            return response;    
        }

        public async Task<ResponseModel<string>> RegistrarPersonaLaboralMasiva(DatosFormatoAsignacionPersonalModel dato, int idUsuario)
        {
            int response = await _usuarioRepository.RegistrarPersonaLaboralMasiva(dato, idUsuario);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con exito");
        }

        public async Task<IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel>> FiltrarAreaPersona(int idArea, string NombrePersona)
        {
            IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel> response = await _usuarioRepository.FiltrarAreaPersona(idArea, NombrePersona);
            return response;
        }

        public async Task<ResponseModel<string>> LiberalPersona(int IdAsignacion)
        {
            int response = await _usuarioRepository.LiberalPersona(IdAsignacion);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Liberado con exito");
        }

        public async Task<ResponseModel<string>> ExportarExcelPersonaAsignacion(DatosFormatoFiltroAsignacionPersona dato)
        {
            if (dato.reporteAsistencia == false && dato.listadoPersonal == false)
                return new ResponseModel<string>(false, "Debe elegir 1 o mas reportes", "");

            IEnumerable<DatosFormatoPersonaAsignacionExportModel> Listar = await _usuarioRepository.ExportarExcelPersonaAsignacion(dato);
            IEnumerable<DatosFormatoListadoPersonalAsignacion> ListaDePersonal = await _usuarioRepository.ExportarExcelListadoPersonal(dato);

            if(Listar.Count() == 0 && ListaDePersonal.Count() == 0)
                return new ResponseModel<string>(false, "No hay información para mostrar en el excel", "");

            ReporteAsignacionPersonal Exporte = new ReporteAsignacionPersonal();
            string reporte = Exporte.GenerarReporte(Listar, ListaDePersonal, dato.reporteAsistencia, dato.listadoPersonal);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }


        public async Task<ResponseModel<AreaPersonalLaboralEntity>> RegistrarEditarArea(int IdArea, string Descripcion)
        {
            AreaPersonalLaboralEntity result = new AreaPersonalLaboralEntity();
            result=await _usuarioRepository.RegistrarEditarArea(IdArea, Descripcion);

            ResponseModel<AreaPersonalLaboralEntity> Respuesta = new ResponseModel<AreaPersonalLaboralEntity>(true, Constante.MESSAGE_SUCCESS, result);

            return Respuesta;
        }

        public async Task<ResponseModel<string>> EliminarAreaProduccion(int IdArea)
        {
             await _usuarioRepository.EliminarAreaProduccion(IdArea);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Eliminado con exito");

            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoListarPersonaTecnica>> ListarPersonaTecnico()
        {
            IEnumerable<DatosFormatoListarPersonaTecnica> response = await _usuarioRepository.ListarPersonaTecnico();

            return response;
        }

        public async Task<IEnumerable<DatosFormatosPersonaPorAreaModel>> ListarPersonaPorArea(int IdArea)
        {
            IEnumerable<DatosFormatosPersonaPorAreaModel> response = await _usuarioRepository.ListarPersonaPorArea(IdArea);

            return response;
        }

        public async Task<IEnumerable<DatosFormatoPersonasAsistencia>> MostrarPersonasAsistencias(string idArea)
        {
            if (string.IsNullOrEmpty(idArea))
                throw new ValidationModelException("Debe Ingresar el Area");

            IEnumerable<DatosFormatoPersonasAsistencia> response = await _usuarioRepository.MostrarPersonasAsistencias(idArea);
            return response;
        }

    }
}
