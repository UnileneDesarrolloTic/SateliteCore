using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using SatelliteCore.Api.Models.Report;
using System;

namespace ReportServices.Contract
{
    public class ReporteCotizacionFactory
    {
        public static CotizacionAbstract ObtenerFormatoCotizacion(int formato, BsonDocument cotizacion)
        {
            switch (formato)
            {
                case 3:

                    Formato3_Model formato3;
                    formato3 = BsonSerializer.Deserialize<Formato3_Model>(cotizacion.AsBsonDocument);

                    return formato3;
                default:
                    throw new EntryPointNotFoundException();
            }
        }
    }
}
