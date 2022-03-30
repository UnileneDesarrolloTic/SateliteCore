using System;

namespace SystemsIntegration.Api.Models.Exceptions
{
    public class NotFoundException : Exception
    {
        public string Mensaje { get; }
        public NotFoundException() : base("Error al validar los datos enviados...") { }

        public NotFoundException(string mensaje) : this()
        {
            Mensaje = mensaje;
        }
    }
}
