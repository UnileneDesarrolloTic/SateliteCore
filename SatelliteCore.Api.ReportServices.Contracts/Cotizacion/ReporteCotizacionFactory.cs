using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System;
using System.Drawing;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public static class ReporteCotizacionFactory
    {
        public static string GenerarReporte(int formato, BsonDocument cotizacion)
        {
            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);
            string reporte;

            switch (formato)
            {
                case Constante.FORMATO3_ESSALUD_ALMENARA:
                    Coti_Formato3_Model formato3 = BsonSerializer.Deserialize<Coti_Formato3_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato3_Report.Exportar(logoUnilene, formato3);
                    return reporte;

                case Constante.FORMATO5_ESSALUD_SABOGAL:
                    Coti_Formato5_Model formato5 = BsonSerializer.Deserialize<Coti_Formato5_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato5_Report.Exportar(logoUnilene, formato5);
                    return reporte;

                default:
                    throw new EntryPointNotFoundException();
            }
        }

        public static CotizacionAbstract ObtenerModeloCotizacion(int formato, BsonDocument cotizacion)
        {

            switch (formato)
            {
                case Constante.FORMATO3_ESSALUD_ALMENARA:                    
                    Coti_Formato3_Model formato1 = BsonSerializer.Deserialize<Coti_Formato3_Model>(cotizacion.AsBsonDocument);
                    return formato1;

                case Constante.FORMATO5_ESSALUD_SABOGAL:
                    Coti_Formato5_Model formato5 = BsonSerializer.Deserialize<Coti_Formato5_Model>(cotizacion.AsBsonDocument);
                    return formato5;

                default:
                    throw new EntryPointNotFoundException();
            }
        }
    }
}
