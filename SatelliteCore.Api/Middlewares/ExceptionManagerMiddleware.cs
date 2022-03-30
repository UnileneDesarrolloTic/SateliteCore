using Microsoft.AspNetCore.Http;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SystemsIntegration.Api.Middlewares
{
    public class ExceptionManagerMiddleware
    {
        private readonly RequestDelegate _next;

        public ExceptionManagerMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task Invoke(HttpContext context)
        {
            try
            {
                await _next(context);
            }
            catch (Exception error)
            {
                HttpResponse response = context.Response;
                response.ContentType = "application/json";

                ResponseModel<List<string>> responseModel =
                        new ResponseModel<List<string>>(false, "", new List<string>());

                switch (error)
                {
                    case ValidationModelException ex:
                        response.StatusCode = (int)HttpStatusCode.BadRequest;
                        responseModel.Message = "Error al validar los datos enviados.";
                        responseModel.Content = ex.Errors;
                        break;
                    case NotFoundException ex:
                        response.StatusCode = (int)HttpStatusCode.NotFound;
                        responseModel.Message = ex.Mensaje;
                        responseModel.Content = null;
                        break;
                    case SqlException ex:
                        response.StatusCode = (int)HttpStatusCode.BadGateway;
                        responseModel.Message = ex.Message; // "A ocurrido un error inesperado en servidores externos"; //
                        break;
                    default:
                        response.StatusCode = (int)HttpStatusCode.InternalServerError;
                        responseModel.Message = error.Message; // "A ocurrido un error inesperado en el servicio";
                        break;
                }

                await response.WriteAsync(JsonSerializer.Serialize(responseModel));
            }

        }
    }
}
