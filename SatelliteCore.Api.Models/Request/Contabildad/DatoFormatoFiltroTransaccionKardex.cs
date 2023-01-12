using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Contabildad
{
    public struct DatoFormatoFiltroTransaccionKardex
    {
        public string Periodo { get; set; }
        public string Tipo { get; set; }
        public int Pagina { get; set; }
        public int RegistrosPorPagina { get; set; }
        public bool CheckCierre { get; set; }
    }
}
