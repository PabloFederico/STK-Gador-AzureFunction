using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Expo.Clases;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;


namespace Expo
{
    public static class ExpoFunction
    {
        [FunctionName("ExpoFunction")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string name = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0)
                .Value;

            //Paramteros Nuevo Elemento
            string _pedido = null;
            string _cliente = null;
            string _urlSitio = null;
            string _factura_url = null;
            string _factura_titulo = null;
            string _tipo = null;

            //Paramteros Edicion
            string _id_editItem = null;


            if (name == null)
            {
                // Get request body
                dynamic data = await req.Content.ReadAsAsync<object>();
                _tipo = data?.Tipo;
                name = data?.name;
                _pedido = data?.pedido;
                _cliente = data?.cliente;
                _urlSitio = data?.urlSitio;
                _factura_titulo = data?.factura_titulo;
                _factura_url = data?.factura_url;
                _id_editItem = data?.id;
                
            }

            Pedidos pedidos = new Pedidos(_urlSitio);
            string responseHTTP = null;

            //NUEVO PEDIDO O EDICION DE UN PEDIDO
            if (_tipo == "NuevoElemento")
            {
                string factura = _pedido;

                if (!pedidos.ExistePedido(factura))
                {
                    pedidos.CrearEstructura(factura, _cliente, _factura_url, _factura_titulo);
                    responseHTTP = pedidos.NombreSitio();
                }
                else
                    responseHTTP = "Factura ya existente";
            }
            else
            {
                int id = Int32.Parse(_id_editItem);
                pedidos.ItemUpdated("N° Pedido", id, _pedido);                       
            }

                      
            return name != null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, responseHTTP);
        }


    }
}
