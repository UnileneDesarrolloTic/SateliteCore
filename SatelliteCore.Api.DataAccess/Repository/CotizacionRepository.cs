using Dapper;
using MongoDB.Bson;
using MongoDB.Driver;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class CotizacionRepository : ICotizacionRepository
    {
        private readonly IAppConfig _appConfig;
        public CotizacionRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<(List<CotizacionEntity>, int)> Listar(DatosListarCotizacionesPaginado datos)
        {
            (List<CotizacionEntity> ListaCertificados, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSpring))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_Cotizacion_ListarCotizaciones", datos, commandType: CommandType.StoredProcedure))
                {
                    result.ListaCertificados = result_db.Read<CotizacionEntity>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }

                connection.Dispose();
            }

            return result;
        }

        public async Task<ObtenerEstructuraFormCotizacionModel> FormatoEstructura(int codFormato)
        {
            ObtenerEstructuraFormCotizacionModel estructura = new ObtenerEstructuraFormCotizacionModel();

            string script = "SELECT CodCampo, Etiqueta, ColumnaResp, TipoDatoTs TipoDato, Requerido, ValorDefecto FROM TBDCamposFormatoCotizacion WHERE IDFormato = @CodFormato AND TipoDetalle = 'U' " +
                            "SELECT CodCampo, Etiqueta, ColumnaResp, TipoDatoTs TipoDato, Requerido, ValorDefecto FROM TBDCamposFormatoCotizacion WHERE IDFormato = @CodFormato AND TipoDetalle = 'D'";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result = await context.QueryMultipleAsync(script, new { codFormato }))
                {
                    estructura.Cabecera = result.Read<EstructuraFormularioCotizacionModel>().ToList();
                    estructura.Detalle = result.Read<EstructuraFormularioCotizacionModel>().ToList();
                }
            }

            return estructura;
        }

        public async Task<(object cabecera, object detalle)> FormatoDatos(int idFormato, string cotizacion)
        {
            (object cabecera, object detalle) datosCotizacion;

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                int formatobd = await context.QueryFirstOrDefaultAsync<int>("SELECT IDFormato FROM TBMFormatoCotizacion WHERE IDFormato=@idFormato", new { idFormato });

                if (formatobd == 0)
                    throw new NotFoundException("El formato enviado no exíste.");

                using (var result = await context.QueryMultipleAsync("usp_Cotizacion_ObtenerDatos", new { idFormato, cotizacion }, commandType: CommandType.StoredProcedure))
                {
                    datosCotizacion.cabecera = result.Read().FirstOrDefault();
                    datosCotizacion.detalle = result.Read().ToList();
                }
            }

            return datosCotizacion;
        }

        public async Task Guardar(string id, string cotizacion, int idFormato, int usuarioSesion)
        {
            string sqlScript = "INSERT INTO TBMRegistroCotizacion (ObjectID, Cotizacion, IdFormato, UsuarioRegistro) VALUES (@id, @cotizacion, @idFormato, @usuarioSesion)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(sqlScript, new { id, cotizacion, idFormato, usuarioSesion });
            }
        }

        public async Task<CotizacionRegistroEntity> ObtenerDatosRegistro(string codReporte)
        {
            CotizacionRegistroEntity response = new CotizacionRegistroEntity();
            string scriptSql = "SELECT ObjectID codigo,Cotizacion,IDFormato,UsuarioRegistro,FechaRegistro,UsuarioModificacion,FechaModificacion FROM TBMRegistroCotizacion WHERE ObjectID=@codReporte";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                response = await context.QueryFirstOrDefaultAsync<CotizacionRegistroEntity>(scriptSql, new { codReporte });
            }

            return response;
        }

        public async Task<IEnumerable<FormatosPorClienteModel>> FormatosPorCliente()
        {
            IEnumerable<FormatosPorClienteModel> listaFormato = new List<FormatosPorClienteModel>();

            string scriptSql = "SELECT a.IdFormato, b.Titulo Formato, a.CodCliente, RTRIM(c.NombreCompleto) Cliente " +
                "FROM TBDFormatoPorClienteCotizacion a INNER JOIN TBMFormatoCotizacion b ON a.IdFormato = b.IDFormato " +
                "INNER JOIN PROD_UNILENE2.dbo.PersonaMast c ON a.CodCliente = c.Persona WHERE c.Estado = 'A'";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                listaFormato = await context.QueryAsync<FormatosPorClienteModel>(scriptSql);
            }

            return listaFormato;
        }

        public async Task<IEnumerable<ReportesGeneradosPorCotizacionModel>> ReportesPorCotizacion(string cotizacion)
        {
            IEnumerable<ReportesGeneradosPorCotizacionModel> reportes = new List<ReportesGeneradosPorCotizacionModel>();

            string scriptSql = "SELECT a.ObjectID Codigo, a.IDFormato IdFormato, b.Titulo Formato, c.Usuario UsuarioRegistro, a.FechaRegistro, d.Usuario UsuarioModificacion, a.FechaModificacion " +
                "FROM TBMRegistroCotizacion a INNER JOIN TBMFormatoCotizacion b ON a.IDFormato = b.IDFormato INNER JOIN TBMUsuario c ON a.UsuarioRegistro = c.CodUsuario " +
                "LEFT JOIN TBMUsuario d ON a.UsuarioModificacion = d.CodUsuario WHERE a.Cotizacion = @cotizacion";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                reportes = await context.QueryAsync<ReportesGeneradosPorCotizacionModel>(scriptSql, new { cotizacion });
            }

            return reportes;
        }


        public async Task<string> Registrar(BsonDocument cotizacion)
        {
            IMongoDatabase integrationContext = new MongoClient(_appConfig.ContextMongoDB).GetDatabase("UnileneReporte");
            IMongoCollection<BsonDocument> _requestModel = integrationContext.GetCollection<BsonDocument>("Cotizacion");

            await _requestModel.InsertOneAsync(cotizacion);

            string idBson = cotizacion.GetValue("_id", "").ToString();
            return idBson;
        }

        public async Task<BsonDocument> ObtenerDatosReporte(string codigoReporte)
        {
            var integrationContext = new MongoClient(_appConfig.ContextMongoDB).GetDatabase("UnileneReporte");
            IMongoCollection<BsonDocument> _requestModel = integrationContext.GetCollection<BsonDocument>("Cotizacion");

            BsonDocument result = await _requestModel.FindAsync(new BsonDocument { { "_id", new ObjectId(codigoReporte) } }).Result.FirstAsync();

            return result;
        }

        public async Task Actualizar(string id, int usuarioSesion, BsonDocument cotizacion)
        {
            CotizacionRegistroEntity datoReporte = await ObtenerDatosRegistro(id);

            if (string.IsNullOrEmpty(datoReporte.Cotizacion))
                throw new NullReferenceException("No se encontró los datos de la cotización"); 

            var integrationContext = new MongoClient(_appConfig.ContextMongoDB).GetDatabase("UnileneReporte");
            IMongoCollection<BsonDocument> _requestModel = integrationContext.GetCollection<BsonDocument>("Cotizacion");

            await _requestModel.ReplaceOneAsync(new BsonDocument { { "_id", new ObjectId(id) } }, cotizacion);

            using (SqlConnection context =  new SqlConnection(_appConfig.contextSatelliteDB) )
            {
                _ = await  context.ExecuteAsync("UPDATE TBMRegistroCotizacion SET FechaModificacion=GETDATE(), UsuarioModificacion=@usuarioSesion WHERE ObjectID=@id", new { usuarioSesion, id });
            }

        }

    }
}
