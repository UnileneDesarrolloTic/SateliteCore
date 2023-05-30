using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public class OrdenServicioModificadosDTO
    {
        public int idTransportista { get; set; }
        public List<OrdenServicioDetalle> ItemsDetalle { get; set; }
    }
    public class OrdenServicioDetalle
    {
        public int Cabecera { get; set; }
        public int Id { get; set; }
        public string Guia { get; set; }
        public DateTime Fecha { get; set; }
        public string Cliente { get; set; }
        public string Direccion { get; set; }
        public string Departamento { get; set; }
        public string Comercial { get; set; }
        public decimal Peso { get; set; }
        public int Bultos { get; set; }
        public string Usuario { get; set; }

        public OrdenServicioDetalle(OrdenServicioDetalle datos, string usuario)
        {
            Cabecera = datos.Cabecera;
            Id = datos.Id;
            Guia = datos.Guia;
            Fecha = datos.Fecha;
            Cliente = datos.Cliente;
            Direccion = datos.Direccion;
            Departamento = datos.Departamento;
            Comercial = datos.Comercial;
            Peso = datos.Peso;
            Bultos = datos.Bultos;
            Usuario = usuario;
        }

        public OrdenServicioDetalle(){}
    }
}
