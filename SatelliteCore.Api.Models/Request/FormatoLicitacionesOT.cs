using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class ListarOpcionesImprimir
    {
        public bool Acta { get; set; }
        public bool Condicion { get; set; }
        public bool Protocolo { get; set; }
        public bool Carta { get; set; }
        public List<FormatoLicitacionesOT> ListaGuias { get; set; }

    }

    public class FormatoLicitacionesOT
    {
        public string GuiasNumero { get; set; }
    }
}
