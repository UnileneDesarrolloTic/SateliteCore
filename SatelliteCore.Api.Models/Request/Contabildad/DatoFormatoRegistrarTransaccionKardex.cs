using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson.Serialization.IdGenerators;
using SatelliteCore.Api.Models.Response.Contabilidad;
using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Contabildad
{

    public class DatoFormatoRegistrarTransaccionKardex
    {
  
        public string Id { get; set; }
        public string Periodo { get; set; }
        public string Tipo { get; set; }
        public bool CheckCierre { get; set; }
        public decimal CCantidadTotal { get; set; }
        public decimal CMontoTotal { get; set; }
        public List<FormatoListadoInformacionTransaccionKardex> InformacionDetalle { get; set; }
    }
}
