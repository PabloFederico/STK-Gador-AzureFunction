using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Expo.Clases
{
    class Constantes
    {
    }
    class Listas
    {
        public const string CarpetaRaiz = "N° Pedido";
        public const string Clientes = "Clientes";
        public const string Subcarpetas = "Subcarpetas";
    }

    class Campos
    {
        //Carpeta Raiz
        public const string Pedido = "Pedido";
        public const string Nombre = "Nombre";
        public const string Factura = "Factura";
        public const string Cliente = "Cliente1";
        public const string Estado = "Estado";
        public const string ID = "ID";

        //Subcarpetas
        public const string sNombre = "Nombre";
        public const string sPedido = "Pedido";
        public const string sID = "ID";

        //Clientes 
        public const string Titulo = "Title";
        public const string cID = "ID";
    }

    class Grupos
    {
        public const string ColaboradoresExportaciones = "Colaboradores Exportaciones";
        public const string PropietariosExportaciones = "Propietarios QA - Exportaciones";
        public const string LectoresPedidos = "Lectores Pedidos";
        public const string LectoresEmbarque = "Lectores Embarque";

    }

    class SubCarpetas
    {
        public const string OrdenCompra = "01 - Orden de Compra";
        public const string ProformaFirmada = "02 - Proforma firmada";
        public const string CertificadoOrigen = "07 - Certificado de Origen";
        public const string DocumentoDeTransporte = "08 - Documento de Transporte";
        public const string Psicotropico = "03 - Psicotrópico de Exportación";
        public const string AvisoSWIFT = "09 - Aviso del Banco - SWIFT";
        public const string DJCera = "04 - DJ CERA firmadas";
        public const string DJInsumosImportados = "05 - DJ Insumos Importados firmada";
        public const string PermisoEmbarque = "06 - Permiso de Embarque";
        public const string Extras = "10 - Extras";
    }
}
