using System.ComponentModel.DataAnnotations;
namespace SatelliteCore.Api.Models.Request.Contabildad
{
    public struct DatoFormatoFiltroTransaccionKardex
    {
        [Required]
        public string Periodo { get; set; }
        [Required]
        public string Tipo { get; set; }
        public bool CheckCierre { get; set; }
        public int Pagina { get; set; }
        public int RegistrosPorPagina { get; set; }

    }
}
