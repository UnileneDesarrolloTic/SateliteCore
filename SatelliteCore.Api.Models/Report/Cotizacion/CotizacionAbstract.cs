using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    public abstract class CotizacionAbstract
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonIgnoreIfNull]
        public string Id { get; set; }
    }
}
