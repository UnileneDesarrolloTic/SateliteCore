﻿using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IProduccionServices
    {
        public Task<List<ProductoArimaModel>> SeguimientoProductosArima(string periodo);
        public Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro);
        public Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla);
        public Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item);
        public Task<bool> MostrarColumnaMP(int usuario);
        public Task<List<CompraMPArimaModel>> SeguimientoCompraMPArima(PronosticoCompraMP dato);
        public Task<ResponseModel<string>> CompraMateriaPrimaExportar(PronosticoCompraMP dato);
        public Task<ResponseModel<FormatoEstructuraLoteEtiquetas>> LoteFabricacionEtiquetas(string NumeroLote);
        public Task<ResponseModel<string>> RegistrarLoteFabricacionEtiquetas(List<DatosEstructuraLoteEtiquetasModel> dato,int idUsuario);
        public Task<IEnumerable<DatoFormatoLoteEstado>> ListarLoteEstado();
        public Task<ResponseModel<string>> ModificarLoteEstado(DatosFormatoRequestLoteEstado dato);
        public Task<DatosFormatoInformacionCalendarioSeguimientoOC> ListarItemOrdenCompra(string Origen, string Anio);
        public Task<DatosFormatoInformacionItemOrdenCompra> BuscarItemOrdenCompra(string Item,string Anio);
        public Task<ResponseModel<string>> ActualizarFechaPrometida(DatosFormatoItemActualizarItemOrdenCompra dato);
        public Task<(object cabecera, object detalle)> VisualizarOrdenCompra(string OrdenCompra);
        public Task<ResponseModel<string>> ActualizarFechaPrometidaMasiva(List<DatosFormatoItemActualizarItemOrdenCompra> dato);
    }
}
