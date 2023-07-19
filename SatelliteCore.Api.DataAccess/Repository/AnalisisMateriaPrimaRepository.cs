using Dapper;
using SatelliteCore.Api.DataAccess.Contracts;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class AnalisisMateriaPrimaRepository : IAnalisisMateriaPrimaRepository
    {
        private readonly IAppConfig _appConfig;

        public AnalisisMateriaPrimaRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ListaAnalisisMateriaPrimaDTO>> ListaOrdenesCompra(string ordenCompra, string codigoAnalisis)
        {
            IEnumerable<ListaAnalisisMateriaPrimaDTO> result = new List<ListaAnalisisMateriaPrimaDTO>();
            string query = "SELECT a.ControlNumero, RTRIM(a.NumeroOrden) NumeroOrden, RTRIM(b.Lote) Analisis, d.DescripcionCompleta Tipo," +
                "RTRIM(b.Item) Item, RTRIM(c.DescripcionLocal) Descripcion,b.CantidadAceptada, b.FechaAprobacion, IIF((c.Linea NOT IN ('N', 'I') OR c.Familia <> 'ST'), 'Otros', CASE WHEN c.SubFamilia IN ('NA','NL','NN','NT','NS', " +
                "'PD','XA','XI','PE','PF','PB','PG','PA','PX','PI','PR','PC','PY','PN','PO','PL','SB','SN','VA','VN','CR','CS','AP','AA','AI') THEN 'Hebra' ELSE 'Otros' END) TipoItem " +
                "FROM WH_ControlCalidad a " +
                    "INNER JOIN WH_ControlCalidadDetalle b ON a.CompaniaSocio = b.CompaniaSocio AND a.ControlNumero = b.ControlNumero " +
                    "INNER JOIN WH_ItemMast c ON b.Item = c.Item " +
                    "INNER JOIN WH_ClaseSubFamilia d ON d.Linea = 'I' AND d.Familia = 'ST' AND c.SubFamilia = d.SubFamilia " +
                 "WHERE a.CompaniaSocio = '01000000' " +
                    "AND a.NumeroOrden = IIF(ISNULL(@OrdenCompra, '') = '', a.NumeroOrden, @OrdenCompra) " +
                    "AND b.Lote = IIF(ISNULL(@codigoAnalisis, '') = '', b.Lote, @codigoAnalisis)";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<ListaAnalisisMateriaPrimaDTO>(query, new { ordenCompra, codigoAnalisis });
            }

            return result;
        }

        public async Task<bool> ValidarSiExisteAnalisisHebra(string ordenCompra, string numeroAnalisis)
        {
            bool exists = false;

            string query = "IF EXISTS (SELECT 1 FROM TBMAnalisisHebra WHERE OrdenCompra = @ordenCompra AND NumeroAnalisis = @numeroAnalisis) SELECT CAST(1 AS BIT) EXISTE " +
                "ELSE SELECT CAST(0 AS BIT) EXISTE";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                exists = await context.QueryFirstOrDefaultAsync<bool>(query, new { ordenCompra, numeroAnalisis });
            }

            return exists;

        }

        public async Task CrearAnalisisHebra(GuardarAnalisisHebraDTO analisis)
        {
            string queryCabecera = "INSERT INTO TBMAnalisisHebra (OrdenCompra, NumeroAnalisis, FechaAnalisis, Certificado, Quimica, Conclusion, " +
                "Observaciones, UsuarioRegistro, FechaRegistro) VALUES(@OrdenCompra, @NumeroAnalisis, @FechaAnalisis, @Certificado, @Quimica, @Conclusion, @Observaciones, @UsuarioRegistro, @FechaRegistro)";

            string queryDetalle = "INSERT INTO TBDAnalisisHebra (OrdenCompra, NumeroAnalisis, Numero, Longitud, Diametro, Tension) " +
                "VALUES(@OrdenCompra, @NumeroAnalisis, @Numero, @Longitud, @Diametro, @Tension)";
            
            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(queryCabecera, analisis.Cabecera);
                await context.ExecuteAsync(queryDetalle, analisis.Detalle);
            }
        }

        public async Task ActualizarAnalisisHebra(GuardarAnalisisHebraDTO analisis)
        {
            string queryCabecera = "UPDATE TBMAnalisisHebra SET FechaAnalisis = @FechaAnalisis, Certificado = @Certificado, " +
                "Quimica = @Quimica, Conclusion = @Conclusion, Observaciones = @Observaciones, UsuarioModificacion = @UsuarioModificacion, " +
                "FechaModificacion = GETDATE() WHERE OrdenCompra = @OrdenCompra AND NumeroAnalisis = @NumeroAnalisis";

            string queryDetalle = "UPDATE TBDAnalisisHebra SET Longitud = @Longitud, Diametro = @Diametro, Tension = @Tension " +
                "WHERE OrdenCompra = @OrdenCompra AND NumeroAnalisis = @NumeroAnalisis AND Numero = @Numero";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                await context.ExecuteAsync(queryCabecera, analisis.Cabecera);
                await context.ExecuteAsync(queryDetalle, analisis.Detalle);
            }
        }

        public async Task<AnalisisHebraDatosGeneralesDTO> DatosGeneralesAnalisisHebra(string ordenCompra, string numeroAnalisis)
        {
            AnalisisHebraDatosGeneralesDTO datos = new AnalisisHebraDatosGeneralesDTO();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync("usp_satelite_AnalisisHebraDatosGenerales", new { ordenCompra, numeroAnalisis }, commandType: CommandType.StoredProcedure);
                datos.Datos = multi.Read<DatosAnalisisHebraDTO>().FirstOrDefault();
                datos.Cabecera = multi.Read<TBMAnalisisHebraEntity>().FirstOrDefault();
                datos.Detalle = multi.Read<TBDAnalisisHebraEntity>().ToList();
            }

            return datos;
        }
        
        public async Task<List<PlantillaDetalleProtocoloDTO>> DatosProtocoloAnalisis(string ordenCompra, string numeroAnalisis)
        {
            IEnumerable<PlantillaDetalleProtocoloDTO> result = new List<PlantillaDetalleProtocoloDTO>();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                result = await context.QueryAsync<PlantillaDetalleProtocoloDTO>("usp_Satelite_ObtenerDatosProtocolo", new { ordenCompra, numeroAnalisis }, commandType: CommandType.StoredProcedure);
            }

            return result.ToList();
        }

        public async Task GuardarDatosProtocoloMateriPrima(List<GuardarProtocoloMateriaPrimaDTO> protocolo)
        {
            GuardarProtocoloMateriaPrimaDTO datos = protocolo.FirstOrDefault();

            string queryDelete = "DELETE WH_ItemProtocolo WHERE Item = @item and NumeroLote = @numeroLote AND ItemNumeroParte = @itemNumeroParte";

            string query = "DECLARE @fechaAnalisis DATETIME " +
                "SELECT @fechaAnalisis = ISNULL(FechaAnalisis, GETDATE()) FROM SatelliteCore.dbo.TBMAnalisisHebra WHERE  NumeroAnalisis = @NumeroLote " +
                "INSERT INTO WH_ItemProtocolo (Item, NumeroLote, ItemNumeroParte, Secuencia, FechaAnalisis, Prueba, Especificacion, Metodologia, Valor, TipoDato, " +
                "Minimo, Maximo, Rechazado, Aprobado, Estado, ConclusionFlag, Comentarios, UltimoUsuario, UltimaFechaModif) " +
                "VALUES(@Item, @NumeroLote, @ItemNumeroParte, @Secuencia, @fechaAnalisis, @Prueba, @Especificacion, @Metodologia, @Valor, @TipoDato, @Minimo, @Maximo, @Rechazado, " +
                "@Aprobado, 'A', @conclusion, @observaciones, @Usuario, GETDATE())";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                await context.ExecuteAsync(queryDelete, new { datos.Item, datos.NumeroLote, datos.ItemNumeroParte });
                await context.ExecuteAsync(query, protocolo);
            }

        }

        public async Task<bool> ValidarRegistroAnalisisHebra(string ordenCompra, string numeroAnalisis)
        {
            bool exists = false;

            string query = "IF EXISTS (SELECT 1 FROM TBMAnalisisHebra WHERE OrdenCompra = @ordenCompra AND NumeroAnalisis = @numeroAnalisis) SELECT CAST(1 AS BIT) EXISTE " +
                "ELSE SELECT CAST(0 AS BIT) EXISTE";

            using (SqlConnection context = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                exists = await context.QueryFirstOrDefaultAsync<bool>(query, new { ordenCompra, numeroAnalisis });
            }

            return exists;
        }

        public async Task<PlantillaProtocoloDTO> DatosReporteProtocolo(string ordenCompra, string numeroAnalisis)
        {
            PlantillaProtocoloDTO datosReporte = new PlantillaProtocoloDTO();

            using (SqlConnection context = new SqlConnection(_appConfig.contextSpring))
            {
                using SqlMapper.GridReader multi = await context.QueryMultipleAsync("usp_Satelite_DatosProtocoloMateriaPrima", new { ordenCompra, numeroAnalisis }, commandType: CommandType.StoredProcedure);
                datosReporte.Cabecera = multi.Read<PlantillaCabeceraProtocoloDTO>().FirstOrDefault();
                datosReporte.Detalle = multi.Read<PlantillaDetalleProtocoloDTO>().ToList();
            }

            return datosReporte;
        }



    }
}
