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
                case Constante.FORMATO_3_ESSALUD_ALMENARA:
                    Formato3_Model formato3 = BsonSerializer.Deserialize<Formato3_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato3_Report.Exportar(logoUnilene, formato3);
                    return reporte;

                case Constante.FORMATO_5_ESSALUD_SABOGAL:
                    Formato5_Model formato5 = BsonSerializer.Deserialize<Formato5_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato5_Report.Exportar(logoUnilene, formato5);
                    return reporte;

                case Constante.FORMATO_60_ESSALUD_PRESTACIONAL_ALMENARA:
                    Formato60_Model formato60 = BsonSerializer.Deserialize<Formato60_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato60_Report.Exportar(logoUnilene, formato60);
                    return reporte;

                case Constante.FORMATO_61_ESSALUD_INCOR:
                    Formato61_Model formato61 = BsonSerializer.Deserialize<Formato61_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato61_Report.Exportar(logoUnilene, formato61);
                    return reporte;

                case Constante.FORMATO_10_ESSALUD_AREQUIPA:
                    Formato10_Model formato10 = BsonSerializer.Deserialize<Formato10_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato10_Report.Exportar(logoUnilene, formato10);
                    return reporte;

                case Constante.FORMATO_65_GENERAL:
                    Formato65_Model formato65 = BsonSerializer.Deserialize<Formato65_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato65_Report.Exportar(logoUnilene, formato65);
                    return reporte;

                default:
                    throw new EntryPointNotFoundException();
            }
        }

        public static CotizacionAbstract ObtenerModeloCotizacion(int formato, BsonDocument cotizacion)
        {

            switch (formato)
            {
                case Constante.FORMATO_3_ESSALUD_ALMENARA:
                    Formato3_Model formato1 = BsonSerializer.Deserialize<Formato3_Model>(cotizacion.AsBsonDocument);
                    return formato1;

                case Constante.FORMATO_5_ESSALUD_SABOGAL:
                    Formato5_Model formato5 = BsonSerializer.Deserialize<Formato5_Model>(cotizacion.AsBsonDocument);
                    return formato5;

                case Constante.FORMATO_60_ESSALUD_PRESTACIONAL_ALMENARA:
                    Formato60_Model formato60 = BsonSerializer.Deserialize<Formato60_Model>(cotizacion.AsBsonDocument);
                    return formato60;

                case Constante.FORMATO_61_ESSALUD_INCOR:
                    Formato61_Model formato61 = BsonSerializer.Deserialize<Formato61_Model>(cotizacion.AsBsonDocument);
                    return formato61; 

                case Constante.FORMATO_10_ESSALUD_AREQUIPA:
                    Formato10_Model formato10 = BsonSerializer.Deserialize<Formato10_Model>(cotizacion.AsBsonDocument);
                    return formato10;

                case Constante.FORMATO_65_GENERAL:
                    Formato65_Model formato65 = BsonSerializer.Deserialize<Formato65_Model>(cotizacion.AsBsonDocument);
                    return formato65;

                default:
                    throw new EntryPointNotFoundException();
            }
        }
    }
}
