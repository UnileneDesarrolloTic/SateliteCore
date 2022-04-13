using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato3_Model : CotizacionAbstract
    {
        // DATOS DEL PROVEEDOR (UNILENE)
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Correo { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_NroCotizacion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Fax { get; set; }
        public string Prov_Vigencia { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public string Prov_DatosAdicionales { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Cargo { get; set; }
        public string Prov_Celular { get; set; }
        public string Prov_Email { get; set; }

        // DATOS DEL AREA SOLICITANTE
        public string Soli_AreaUsuario { get; set; }
        public string Soli_Dependencia { get; set; }
        public string Soli_Contacto { get; set; }
        public string Soli_Correo { get; set; }
        public string Soli_Telefono { get; set; }

        // PIE DE PAGINA
        public string Foot_FormaPago { get; set; }
        public string Foot_Observaciones { get; set; }
        public string Foot_Importante { get; set; }
        public string Foot_Firma { get; set; }

        public decimal Total { get; set; }

        public List<Formato3_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato3_Detalle
    {
        public int NroItem { get; set; }
        public string CodigoSAP { get; set; }
        public string Denominacion { get; set; }
        public string UndMedida { get; set; }
        public decimal CantidadTotal { get; set; }
        public string Marca { get; set; }
        public string PaisProcedencia { get; set; }
        public string Presentacion { get; set; }
        public string Modelo { get; set; }
        public string RnpVigente { get; set; }
        public string VigenciaMinima { get; set; }
        public string CumpleDenominacionItem { get; set; }
        public string CartaPresentacion { get; set; }
        public string RegistroSanitario { get; set; }
        public string CertificadoAnalisis { get; set; }
        public string CumpleEspecificacionTec { get; set; }
        public string PracticaAlmacenamiento { get; set; }
        public string PracticaManufactura { get; set; }
        public string CapacidadAtencion { get; set; }
        public string PlazoEntrega { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal ValorTotal { get; set; }
    }
}
