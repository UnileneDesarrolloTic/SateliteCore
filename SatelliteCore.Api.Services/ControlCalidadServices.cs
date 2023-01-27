using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.Comercial;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.ReportServices.Contracts.ControlCalidad;

namespace SatelliteCore.Api.Services
{
    public class ControlCalidadServices : IControlCalidadServices
    {
        private readonly IControlCalidadRepository _controlCalidadRepository;
        private readonly ICommonRepository _commonRepository;


        public ControlCalidadServices(IControlCalidadRepository controlCalidadRepository, ICommonRepository commonRepository)
        {
            _controlCalidadRepository = controlCalidadRepository;
            _commonRepository = commonRepository;
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

        public async Task<DatosFormatoListarOrdenFabricacionModel> ObtenerInformacionLote(string NumeroLote)
        {
            return await _controlCalidadRepository.ObtenerInformacionLote(NumeroLote);
        }

        public async Task<IEnumerable<DatosFormatoListarTransaccion>> ListarTransaccionItem(string NumeroLote, string codAlmacen)
        {
            return await _controlCalidadRepository.ListarTransaccionItem(NumeroLote, codAlmacen);
            
        }

        public async Task<ResponseModel<string>> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato,int idUsuario)
        {
            int reponse = await _controlCalidadRepository.RegistrarLoteNumeroCaja(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con éxito");
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoKardexInternoGCM>> ListarKardexInternoNumeroLote(string NumeroLote)
        {
            return await _controlCalidadRepository.ListarKardexInternoNumeroLote(NumeroLote);
        }

        public async Task<ResponseModel<string>> ActualizarKardexInternoGCM(int idKardex, string comentarios, int idUsuario)
        {
            int reponse = await _controlCalidadRepository.ActualizarKardexInternoGCM(idKardex, comentarios, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato, int idUsuario)
        {
            int reponse = await _controlCalidadRepository.RegistrarKardexInternoGCM(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> ExportarOrdenFabricacionCaja(string anioProduccion)
        {
            IEnumerable<FormatoEstructuraObtenerOrdenFabricacion> listaOrdenFabricacionCaja =  await _controlCalidadRepository.ExportarOrdenFabricacionCaja(anioProduccion);
            ReporteOrdenFabricacionCaja ExporteOrdenFabricacionCaja = new ReporteOrdenFabricacionCaja();
            string reporte = ExporteOrdenFabricacionCaja.GenerarReporteCaja(listaOrdenFabricacionCaja);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
            return Respuesta;
        }


        public async Task<IEnumerable<DatosFormatosListarControlLotes>> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato)
        {
            IEnumerable<DatosFormatosListarControlLotes> lista = await _controlCalidadRepository.ListarControlLotes(dato);
            return lista;
        }

        public async Task<ResponseModel<string>> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato)
        {
            int reponse = await _controlCalidadRepository.ActualizarControlLotes(dato);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoTablaNumerodeParte>> ListarMaestroNumeroParte(string Grupo, string Tabla, string Usuario)
        {
            IEnumerable<DatosFormatoTablaNumerodeParte> result = new List<DatosFormatoTablaNumerodeParte>();
            IEnumerable <DatosFormatoTablaNumerodeParte> Respuesta = await _controlCalidadRepository.ListarMaestroNumeroParte(Grupo,Tabla);
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_BUSCAR_NUMERO_DE_PARTE;
            evento.Usuario = Usuario;
            await _commonRepository.RegistroLogEvento(evento);

            foreach (DatosFormatoTablaNumerodeParte valor in Respuesta)
            {
                DatosFormatoTablaNumerodeParte myObj = new DatosFormatoTablaNumerodeParte();
                myObj.Grupo = valor.Grupo.Trim();
                myObj.NombreGrupo = valor.NombreGrupo.Trim();
                myObj.CodigoTabla = valor.CodigoTabla.Trim();
                myObj.NombreTabla = valor.NombreTabla.Trim();
                myObj.Codigo = valor.Codigo.Trim();
                myObj.DescripcionLocal = valor.DescripcionLocal.Trim();
                myObj.Longitud = valor.Longitud;
                myObj.Estado = valor.Estado;
                myObj.UltimaFechaModif = valor.UltimaFechaModif;
                myObj.UltimoUsuario = valor.UltimoUsuario;
                result=result.Concat(new[] { myObj });
            }
            
            return result;
        }

        public async Task<IEnumerable<DatosFormatoTablaAbributoModel>> ListarAtributos()
        {
            IEnumerable<DatosFormatoTablaAbributoModel> response = await _controlCalidadRepository.ListarAtributos();
            return response;
        }

        public async Task<IEnumerable<DatosFormatoTablaDescripcionModel>> ListarDescripcion(string Marca, string Hebra,string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_BUSCAR_DESCRIPCION;
            evento.Usuario = Usuario;

            IEnumerable<DatosFormatoTablaDescripcionModel> response = await _controlCalidadRepository.ListarDescripcion(Marca, Hebra);
            await _commonRepository.RegistroLogEvento(evento);
            return response;
        }
        public async Task<IEnumerable<DatosFormatoTablaLeyendaModel>> ListarLeyenda(string Marca, string Hebra,string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_BUSCAR_LEYENDA;
            evento.Usuario = Usuario;

            IEnumerable<DatosFormatoTablaLeyendaModel> response = await _controlCalidadRepository.ListarLeyenda(Marca, Hebra);
            await _commonRepository.RegistroLogEvento(evento);
            return response;
        }

        public async Task<IEnumerable<DatosFormatoTablaPruebasModel>> ListarTablaPrueba(string Metodologia,string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_BUSCAR_PRUEBAS;
            evento.Usuario = Usuario;

            IEnumerable<DatosFormatoTablaPruebasModel> response = await _controlCalidadRepository.ListarTablaPrueba(Metodologia);
            await _commonRepository.RegistroLogEvento(evento);

            return response;
        }

        public async Task<IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel>> ListarObtenerAgujasDescripcionNuevo()
        {
            IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel> response = await _controlCalidadRepository.ListarObtenerAgujasDescripcionNuevo();
            return response;
        }

 
        public async Task<ResponseModel<string>> NuevoDescripcionDT(DatosFormatoActualizacionDescripcionModel dato, string idUsuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_GUARDAR_DESCRIPCION;
            evento.Usuario = idUsuario;

            await _commonRepository.RegistroLogEvento(evento);
            await _controlCalidadRepository.NuevoDescripcionDT(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> EliminarDescripcionDT(string IdDescripcion,string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_ELIMINAR_DESCRIPCION;
            evento.Usuario = Usuario;
            evento.Opcional = IdDescripcion;

            await _controlCalidadRepository.EliminarDescripcionDT(IdDescripcion);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Eliminado con éxito");
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoObtenerAgujasDescripcionModel>> ListarObtenerAgujasDescripcionActualizar(string IdDescripcion)
        {
            IEnumerable<DatosFormatoObtenerAgujasDescripcionModel> response = await _controlCalidadRepository.ListarObtenerAgujasDescripcionActualizar(IdDescripcion);
            return response;
        }

        public async Task<ResponseModel<string>> ActualizarDescripcionDT(DatosFormatoActualizacionDescripcionModel dato, string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_ACTUALIZAR_DESCRIPCION;
            evento.Usuario = Usuario;

            await _controlCalidadRepository.ActualizarDescripcionDT(dato, Usuario);
            await _commonRepository.RegistroLogEvento(evento);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> RegistrarActualizarLeyendaDT(DatosFormatoLeyendaDTModel dato, string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = dato.IdLeyenda == 0 ? ConstanteLog.AP_GUARDAR_LEYENDA : ConstanteLog.AP_ACTUALIZAR_LEYENDA;
            evento.Usuario = Usuario;
            evento.Opcional = dato.RegistroSanitario;

            await _controlCalidadRepository.RegistrarActualizarLeyendaDT(dato, Usuario);
            await _commonRepository.RegistroLogEvento(evento);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, dato.IdLeyenda==0 ? "Registrar con éxito" : "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> EliminarLeyendaDT(string IdLeyenda,string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.AP_ELIMINAR_LEYENDA;
            evento.Usuario = Usuario;
            evento.Opcional = IdLeyenda;

            await _controlCalidadRepository.EliminarLeyendaDT(IdLeyenda);
            await _commonRepository.RegistroLogEvento(evento);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Eliminado con éxito");
            return Respuesta;
        }


        public async Task<ResponseModel<string>> RegistrarActualizarPruebaDT(DatosFormatoNuevoPruebaModel dato, string idUsuario)
        {
            await _controlCalidadRepository.RegistrarActualizarPruebaDT(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Actualizacion con éxito");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> EliminarPruebaDT(string IdPrueba)
        {
            await _controlCalidadRepository.EliminarPruebaDT(IdPrueba);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Eliminar con éxito");
            return Respuesta;
        }


        public async Task<ResponseModel<DatosFormatoNumeroLoteProtocoloModel>> BuscarNumeroLoteProtocolo(string NumeroLote, string Idioma)
        {
            DatosFormatoNumeroLoteProtocoloModel result = new DatosFormatoNumeroLoteProtocoloModel();
            result= await _controlCalidadRepository.BuscarNumeroLoteProtocolo(NumeroLote, Idioma);
            ResponseModel<DatosFormatoNumeroLoteProtocoloModel> Respuesta = new ResponseModel<DatosFormatoNumeroLoteProtocoloModel>(true, Constante.MESSAGE_SUCCESS, result);
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatosDatoListarPruebaProtocolo>> BuscarPruebaFormatoProtocolo(string NumeroLote, string  NumeroParte, string Idioma)
        {
            IEnumerable<DatosFormatosDatoListarPruebaProtocolo> Respuesta = await _controlCalidadRepository.BuscarPruebaFormatoProtocolo(NumeroLote, NumeroParte, Idioma);
            return Respuesta;
        }

        public async Task<ResponseModel<string>> RegistrarControlProcesoProtocolo(DatosFormatoControlProcesosProtocoloModel dato, string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.FP_GUARDAR_CONTROL_DE_PROCESOS_INTERNO;
            evento.Usuario = Usuario;
            evento.Opcional = dato.Numerolote;

            await _controlCalidadRepository.RegistrarControlProcesoProtocolo(dato, Usuario);
            await _commonRepository.RegistroLogEvento(evento);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con existo");
            return Respuesta;
        }
        public async Task<ResponseModel<string>> RegistrarControlPTProtocolo(DatosFormatoControlProductoTermino dato, string Usuario)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.FP_GUARDAR_CONTROL_DE_PRUEBA_PT;
            evento.Usuario = Usuario;
            evento.Opcional = dato.Numerolote;

            await _controlCalidadRepository.RegistrarControlPTProtocolo(dato, Usuario);
            await _commonRepository.RegistroLogEvento(evento);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con existo");
            return Respuesta;
        }
        public async Task<ResponseModel<string>> RegistrarPruebasEfectuadasProtocolo(DatosFormatoPruebasEfectuasProtocolos dato, string idUsuario)
        {
            await _controlCalidadRepository.RegistrarPruebasEfectuadasProtocolo(dato, idUsuario);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con existo");
            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoInformacionResultadoProtocolo>> BuscarInformacionResultadoProtocolo(string NumeroLote)
        {
            IEnumerable<DatosFormatoInformacionResultadoProtocolo> listado = await _controlCalidadRepository.BuscarInformacionResultadoProtocolo(NumeroLote);
             return listado;
        }

        public async Task<ResponseModel<string>> InsertarCabeceraFormatoProtocolo(DatosFormatoCabeceraFormatoProtocolo dato, string UsuarioSesion)
        {
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.FP_GUARDAR_PRUEBAS_EFECTUADAS;
            evento.Usuario = UsuarioSesion;
            evento.Opcional = dato.NumeroLote;

            await _controlCalidadRepository.InsertarCabeceraFormatoProtocolo(dato,UsuarioSesion);
            await _commonRepository.RegistroLogEvento(evento);
            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con existo");
            return Respuesta;
        }

        public async Task<ResponseModel<string>> ImprimirControlProcesoInterno(string NumeroLote, string UsuarioSesion)
        {
            LogTrazaEvento evento = new LogTrazaEvento();

            DatosFormatoNumeroLoteProtocoloModel Cabecera = new DatosFormatoNumeroLoteProtocoloModel();
            IEnumerable<DatosFormatoInformacionResultadoProtocolo> listado = await _controlCalidadRepository.ImprimirControlProceso(NumeroLote);
            if(listado.Count()==0)
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No hay información registrada");

            Cabecera = await _controlCalidadRepository.BuscarNumeroLoteProtocolo(NumeroLote,"1");
            ControlProcesoInterno ExporteControlProcesoInterno = new ControlProcesoInterno();
            string reporte = ExporteControlProcesoInterno.ReporteControlProcesoInterno(listado,Cabecera);

            evento.IdEvento = ConstanteLog.FP_IMPRIMIR_CONTROL_DE_PROCESOS_INTERNO;
            evento.Usuario = UsuarioSesion;
            evento.Opcional = NumeroLote;
            await _commonRepository.RegistroLogEvento(evento);


            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
            return Respuesta;
        }


        public async Task<ResponseModel<string>> ImprimirControlPruebas(string NumeroLote,string UsuarioSesion)
        {
            DatosFormatoNumeroLoteProtocoloModel Cabecera = new DatosFormatoNumeroLoteProtocoloModel();
            LogTrazaEvento evento = new LogTrazaEvento();

            IEnumerable<DatosFormatoInformacionResultadoProtocolo> listado = await _controlCalidadRepository.ImprimirControlProceso(NumeroLote);
            Cabecera = await _controlCalidadRepository.BuscarNumeroLoteProtocolo(NumeroLote,"1");
            ControldePruebas ExporteControldePruebas = new ControldePruebas();
            string reporte = ExporteControldePruebas.ReporteControldePruebas(listado, Cabecera);
            evento.IdEvento = ConstanteLog.FP_IMPRIMIR_CONTROL_DE_PRUEBA_PT;
            evento.Usuario = UsuarioSesion;
            evento.Opcional = NumeroLote;
            await _commonRepository.RegistroLogEvento(evento);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
            return Respuesta;
        }

        public async Task<ResponseModel<string>> ImprimirDocumentoProtocolo(string NumeroLote, bool Opcion, string Idioma,string UsuarioSesion)
        {
            string reporte ="";
            DatosFormatoNumeroLoteProtocoloModel Cabecera = new DatosFormatoNumeroLoteProtocoloModel();
            ParametroMastEntity datosPiePagina = new ParametroMastEntity();
            LogTrazaEvento evento = new LogTrazaEvento();
            evento.IdEvento = ConstanteLog.FP_IMPRIMIR_PRUEBAS_EFECTUADAS;
            evento.Usuario = UsuarioSesion;
            evento.Opcional = "Lote: " + NumeroLote +" Idioma: "+ Idioma  + " Firma: " + Opcion.ToString();
            await _commonRepository.RegistroLogEvento(evento);

            IEnumerable<DatosFormatoProtocoloPruebaModel> listado = await _controlCalidadRepository.ImprimirDocumentoProtocolo(NumeroLote, Idioma);

            if (listado.Count() == 0)
                return new ResponseModel<string>(false, Constante.MESSAGE_SUCCESS, "No hay Pruebas Efectuadas para ese lote");


            Cabecera = await _controlCalidadRepository.BuscarNumeroLoteProtocolo(NumeroLote, Idioma);
            datosPiePagina = await _controlCalidadRepository.ProtocoloRevisionTerminado();

            FormatoPruebaProtocolo ExporteFormatoPrueba = new FormatoPruebaProtocolo();
            
            if (Idioma=="1")
                reporte = ExporteFormatoPrueba.ReporteFormatoPruebaProtocoloEspaniol(listado, Cabecera, Opcion , datosPiePagina);
            else
                reporte = ExporteFormatoPrueba.ReporteFormatoPruebaProtocoloIngles(listado, Cabecera, Opcion, datosPiePagina);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);
            return Respuesta;
        }




    }
}
