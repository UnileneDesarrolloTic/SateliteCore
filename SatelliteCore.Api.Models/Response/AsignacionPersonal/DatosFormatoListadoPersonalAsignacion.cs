using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.AsignacionPersonal
{
    public struct DatosFormatoListadoPersonalAsignacion
    {
        public DateTime Fecha { get; set; }
        public string ClasificacionGerencia { get; set; }
        public int Persona { get; set; }
        public int IdAsignacion { get; set; }
        public string NombreCompleto { get; set; }
        public string Horario { get; set; }
        public string NombreArea { get; set; }
        public DateTime? FechaAsignacion { get; set; }
        public DateTime? FechaReAsignacion { get; set; }
        public string Estado { get; set; }
        public string Ingreso { get; set; }
        public string Salida { get; set; }
        public int Hreal { get; set; }
        public int NumDia { get; set; }
        public int Tardanza { get; set; }
        public int Hextra { get; set; }
        public int Justificaciones { get; set; }
        public int Vacaciones { get; set; }

    }
}
