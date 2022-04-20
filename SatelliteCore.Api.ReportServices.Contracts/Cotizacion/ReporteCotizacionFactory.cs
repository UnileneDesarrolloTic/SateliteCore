using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.Models.Report.Cotizacion;
using System.Drawing;
using System.IO;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.ReportServices.Contracts.Cotizacion
{
    public static class ReporteCotizacionFactory
    {
        public static string GenerarReporte(int formato, BsonDocument cotizacion)
        {
            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);

            string rutaFirma = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Firma_Cotizaciones.jpeg");
            Image firma = Image.FromFile(rutaFirma);

            string rutaEsSalud = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_EsSalud.png");
            Image logoEsSalud = Image.FromFile(rutaEsSalud);

            string reporte;

            switch (formato)
            {
                case Constante.FORMATO_3_ESSALUD_ALMENARA:
                    Formato3_Model formato3 = BsonSerializer.Deserialize<Formato3_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato3_Report.Exportar(firma, logoUnilene, formato3);
                    return reporte;

                case Constante.FORMATO_5_ESSALUD_SABOGAL:
                    Formato5_Model formato5 = BsonSerializer.Deserialize<Formato5_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato5_Report.Exportar(logoUnilene, formato5);
                    return reporte;

                case Constante.FORMATO_60_ESSALUD_PRESTACIONAL_ALMENARA:
                    Formato60_Model formato60 = BsonSerializer.Deserialize<Formato60_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato60_Report.Exportar(firma, logoUnilene, formato60);
                    return reporte;

                case Constante.FORMATO_61_ESSALUD_INCOR:
                    Formato61_Model formato61 = BsonSerializer.Deserialize<Formato61_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato61_Report.Exportar(firma, logoUnilene, formato61);
                    return reporte;

                case Constante.FORMATO_10_ESSALUD_AREQUIPA:
                    Formato10_Model formato10 = BsonSerializer.Deserialize<Formato10_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato10_Report.Exportar(logoUnilene, formato10);
                    return reporte;

                case Constante.FORMATO_65_GENERAL:
                    Formato65_Model formato65 = BsonSerializer.Deserialize<Formato65_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato65_Report.Exportar(logoUnilene, formato65);
                    return reporte;

                case Constante.FORMATO_4_ESSALUD_REBAGLIATI_F2:
                    Coti_Formato4_Model formato4 = BsonSerializer.Deserialize<Coti_Formato4_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato4_Report.Exportar(logoEsSalud, formato4);
                    return reporte;

                case Constante.FORMATO_62_ESSALUD_REBAGLIATI_F1:
                    Coti_Formato62_Model formato62 = BsonSerializer.Deserialize<Coti_Formato62_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato62_Report.Exportar(firma, logoEsSalud, formato62);
                    return reporte;

                case Constante.FORMATO_13_ESSALUD_CUSCO_F1:
                    Coti_Formato13_Model formato13 = BsonSerializer.Deserialize<Coti_Formato13_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato13_Report.Exportar(firma, logoUnilene, formato13);
                    return reporte;

                case Constante.FORMATO_21_ESSALUD_LAMBAYEQUE:
                    Coti_Formato21_Model formato21 = BsonSerializer.Deserialize<Coti_Formato21_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato21_Report.Exportar(logoUnilene, formato21);
                    return reporte;

                case Constante.FORMATO_24_ESSALUD_MOQUEGUA:
                    Coti_Formato24_Model formato24 = BsonSerializer.Deserialize<Coti_Formato24_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato24_Report.Exportar(logoUnilene, formato24);
                    return reporte;

                case Constante.FORMATO_63_INS_NACIONAL_NINIO_BRENIA:
                    Coti_Formato63_Model formato63 = BsonSerializer.Deserialize<Coti_Formato63_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato63_Report.Exportar(logoUnilene, formato63);
                    return reporte;

                case Constante.FORMATO_19_ESSALUD_JUNIN:
                    Coti_Formato19_Model formato19 = BsonSerializer.Deserialize<Coti_Formato19_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato19_Report.Exportar(logoUnilene, formato19);
                    return reporte;

                case Constante.FORMATO_64_HOSPITAL_SAN_BARTOLOME:
                    Coti_Formato64_Model formato64 = BsonSerializer.Deserialize<Coti_Formato64_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato64_Report.Exportar(logoUnilene, formato64);
                    return reporte;

                case Constante.FORMATO_22_ESSALUD_LORETO_F1:
                    Coti_Formato22_Model formato22 = BsonSerializer.Deserialize<Coti_Formato22_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato22_Report.Exportar(logoUnilene, formato22);
                    return reporte;

                case Constante.FORMATO_17_ESSALUD_ICA_F1:
                    Coti_Formato17_Model formato17 = BsonSerializer.Deserialize<Coti_Formato17_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato17_Report.Exportar(logoUnilene, formato17);
                    return reporte;

                case Constante.FORMATO_18_ESSALUD_JULIACA_F1:
                    Coti_Formato18_Model formato18 = BsonSerializer.Deserialize<Coti_Formato18_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato18_Report.Exportar(logoUnilene, formato18);
                    return reporte;

                case Constante.FORMATO_31_ESSALUD_TUMBES:
                    Coti_Formato31_Model formato31 = BsonSerializer.Deserialize<Coti_Formato31_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato31_Report.Exportar(logoUnilene, formato31);
                    return reporte;

                case Constante.FORMATO_66_ESSALUD_CUSCO_F2:
                    Formato66_Model formato66 = BsonSerializer.Deserialize<Formato66_Model>(cotizacion.AsBsonDocument);
                    reporte = Formato66_Report.Exportar(logoUnilene, formato66);
                    return reporte;

                default:
                    throw new NotFoundException("No se puedo encontrar el formato solicitado");
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

                case Constante.FORMATO_4_ESSALUD_REBAGLIATI_F2:
                    Coti_Formato4_Model formato4 = BsonSerializer.Deserialize<Coti_Formato4_Model>(cotizacion.AsBsonDocument);
                    return formato4; 

                case Constante.FORMATO_62_ESSALUD_REBAGLIATI_F1:
                    Coti_Formato62_Model formato62 = BsonSerializer.Deserialize<Coti_Formato62_Model>(cotizacion.AsBsonDocument);
                    return formato62;

                case Constante.FORMATO_13_ESSALUD_CUSCO_F1:
                    Coti_Formato13_Model formato13 = BsonSerializer.Deserialize<Coti_Formato13_Model>(cotizacion.AsBsonDocument);
                    return formato13;

                case Constante.FORMATO_21_ESSALUD_LAMBAYEQUE:
                    Coti_Formato21_Model formato21 = BsonSerializer.Deserialize<Coti_Formato21_Model>(cotizacion.AsBsonDocument);
                    return formato21;

                case Constante.FORMATO_24_ESSALUD_MOQUEGUA:
                    Coti_Formato24_Model formato24 = BsonSerializer.Deserialize<Coti_Formato24_Model>(cotizacion.AsBsonDocument);
                    return formato24;

                case Constante.FORMATO_63_INS_NACIONAL_NINIO_BRENIA:
                    Coti_Formato63_Model formato63 = BsonSerializer.Deserialize<Coti_Formato63_Model>(cotizacion.AsBsonDocument);
                    return formato63;

                case Constante.FORMATO_19_ESSALUD_JUNIN:
                    Coti_Formato19_Model formato19 = BsonSerializer.Deserialize<Coti_Formato19_Model>(cotizacion.AsBsonDocument);
                    return formato19;

                case Constante.FORMATO_64_HOSPITAL_SAN_BARTOLOME:
                    Coti_Formato64_Model formato64 = BsonSerializer.Deserialize<Coti_Formato64_Model>(cotizacion.AsBsonDocument);
                    return formato64;

                case Constante.FORMATO_22_ESSALUD_LORETO_F1:
                    Coti_Formato22_Model formato22 = BsonSerializer.Deserialize<Coti_Formato22_Model>(cotizacion.AsBsonDocument);
                    return formato22;

                case Constante.FORMATO_17_ESSALUD_ICA_F1:
                    Coti_Formato17_Model formato17 = BsonSerializer.Deserialize<Coti_Formato17_Model>(cotizacion.AsBsonDocument);
                    return formato17;

                case Constante.FORMATO_18_ESSALUD_JULIACA_F1:
                    Coti_Formato18_Model formato18 = BsonSerializer.Deserialize<Coti_Formato18_Model>(cotizacion.AsBsonDocument);
                    return formato18;


                case Constante.FORMATO_31_ESSALUD_TUMBES:
                    Coti_Formato31_Model formato31 = BsonSerializer.Deserialize<Coti_Formato31_Model>(cotizacion.AsBsonDocument);
                    return formato31;


                case Constante.FORMATO_66_ESSALUD_CUSCO_F2:
                    Formato66_Model formato66 = BsonSerializer.Deserialize<Formato66_Model>(cotizacion.AsBsonDocument);
                    return formato66;

                default:
                    throw new NotFoundException("No se puedo encontrar el formato solicitado");
            }
        }
    }
}
