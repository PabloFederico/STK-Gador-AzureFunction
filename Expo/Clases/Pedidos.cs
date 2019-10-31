using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System;


namespace Expo.Clases
{
    class Pedidos
    {
        private string url;
        private ClientContext context;

        public Pedidos(string urlSitio)
        {
            this.url = urlSitio;
            this.context = new ClientContext(url);

            string pass = System.Environment.GetEnvironmentVariable("password");
            string USER = System.Environment.GetEnvironmentVariable("user");
          
            var passWord = new System.Security.SecureString();
            foreach (var c in pass) passWord.AppendChar(c);
            this.context.Credentials = new SharePointOnlineCredentials(USER, passWord);
        }

        public void CrearEstructura(string nombreCarpeta, string cliente, string factura, string facturatitulo)
        {
            ClientContext clientContext = this.context;
            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQuery();

            RoleDefinition roleDefinitionViewOnly = null;

            //Obtenemos los distintos grupos;
            Group ColaboradoresExpo = web.SiteGroups.GetByName(Grupos.ColaboradoresExportaciones);
            Group LectoresEmb = web.SiteGroups.GetByName(Grupos.LectoresEmbarque);
            Group LectoresPed = web.SiteGroups.GetByName(Grupos.LectoresPedidos);
            Group PropExp = web.SiteGroups.GetByName(Grupos.PropietariosExportaciones);

            List raiz = web.Lists.GetByTitle(Listas.CarpetaRaiz);
            clientContext.Load(raiz, rt => rt.RootFolder.ServerRelativeUrl);
            clientContext.ExecuteQuery();

            if (raiz != null)
            {
                /***********CREA CARPETA DE PEDIDO***********/
                ListItemCreationInformation list_info = new ListItemCreationInformation();
                list_info.FolderUrl = raiz.RootFolder.ServerRelativeUrl;
                list_info.UnderlyingObjectType = FileSystemObjectType.Folder;
                list_info.LeafName = nombreCarpeta;

                ListItem pedido = raiz.AddItem(list_info);

                //Setea el CT a la carpeta pedido
                ContentTypeCollection ctc = web.ContentTypes;
                clientContext.Load(ctc);
                clientContext.ExecuteQuery();
                ContentType contentType = Enumerable.FirstOrDefault(ctc, ct => ct.Name == "CT.Raiz");
                pedido["ContentTypeId"] = contentType.Id;
                pedido[Campos.Nombre] = nombreCarpeta;
                pedido[Campos.Pedido] = Int64.Parse(nombreCarpeta);
                pedido[Campos.Estado] = "";
                pedido.Update();


                //Campo Cliente  
                if (!string.IsNullOrEmpty(cliente))
                {
                    ListItem itemCliente = this.ObtenerClienteParaLookup(cliente);
                    clientContext.Load(itemCliente);
                    clientContext.ExecuteQuery();

                    if (itemCliente != null)
                    {
                        FieldLookupValue lv = new FieldLookupValue();
                        lv.LookupId = itemCliente.Id;
                        pedido[Campos.Cliente] = lv;

                    }
                }

                //Campo factura
                if (!string.IsNullOrEmpty(factura))
                {
                    FieldUrlValue campoFactura = new FieldUrlValue();
                    campoFactura.Description = string.IsNullOrEmpty(facturatitulo) ? "FACTURA" : facturatitulo;
                    campoFactura.Url = factura;
                    pedido[Campos.Factura] = campoFactura;
                    
                }

                pedido.Update();
                clientContext.ExecuteQuery();

                //Asigna los permisos a la carpeta
                AsignarPermisos(web, pedido, ColaboradoresExpo, RoleType.Contributor);
                AsignarPermisos(web, pedido, PropExp, RoleType.Contributor);
                AsignarPermisos(web, pedido, LectoresPed, RoleType.Reader);
                AsignarPermisos(web, pedido, LectoresEmb, RoleType.Reader);

                //pedido.UpdateOverwriteVersion();


                /*********************************************SUBCARPETAS***************************************************************/
                List<string> carpetas = ObtenerNombresSubcarpetas(web.Url);
                
                try
                {
                     roleDefinitionViewOnly = web.RoleDefinitions.Cast<RoleDefinition>().FirstOrDefault(r => r.Name.ToUpper() == "VISTA SÓLO");
                }
                catch
                {
                }

                foreach (string carpeta in carpetas)
                {
                    //Crea la subcarpeta
                    ListItemCreationInformation subfolder_info = new ListItemCreationInformation();
                    subfolder_info.FolderUrl = raiz.RootFolder.ServerRelativeUrl + "/" + nombreCarpeta;
                    subfolder_info.UnderlyingObjectType = FileSystemObjectType.Folder;
                    subfolder_info.LeafName = carpeta;
                    ListItem sfolder = raiz.AddItem(subfolder_info);


                    //Setea el CT a la subcarpeta
                    ContentTypeCollection ct_sub = web.ContentTypes;
                    clientContext.Load(ct_sub);
                    clientContext.ExecuteQuery();
                    ContentType contentType_sub = Enumerable.FirstOrDefault(ct_sub, ct => ct.Name == "CT.Sub");
                    sfolder["ContentTypeId"] = contentType_sub.Id;
                    sfolder.Update(); //Para obtener la url

                    //Setea campos
                    sfolder[Campos.Nombre] = carpeta;
                    sfolder[Campos.Pedido] = Int64.Parse(nombreCarpeta);
                    sfolder.Update();                   

                    /*
                    if (!string.IsNullOrEmpty(factura))
                    {
                        FieldUrlValue campoFac = new FieldUrlValue();
                        campoFac.Description = string.IsNullOrEmpty(facturatitulo) ? "FACTURA" : facturatitulo;
                        campoFac.Url = factura;
                        sfolder[Campos.Factura] = campoFac;
                    }
              
                    if (!string.IsNullOrEmpty(cliente))
                    {
                        ListItem itemclie = this.ObtenerClienteParaLookup(cliente);
                        clientContext.Load(itemclie);
                        clientContext.ExecuteQuery();

                        FieldLookupValue lookup = new FieldLookupValue();
                        lookup.LookupId = itemclie.Id;
                        sfolder[Campos.Cliente] = lookup;
                        sfolder.Update();
                        clientContext.ExecuteQuery();


                    }*/

                    clientContext.ExecuteQuery();

                    switch (carpeta)
                    {
                        case SubCarpetas.PermisoEmbarque:
                            this.AsignarPermisos(web, sfolder, ColaboradoresExpo, RoleType.Contributor);
                            this.AsignarPermisos(web, sfolder, PropExp, RoleType.Contributor);
                            if (roleDefinitionViewOnly != null)
                                this.AsignarPermisosPorRoleDefinition(web, sfolder, LectoresEmb, roleDefinitionViewOnly);
                            else
                            //this.AsignarPermisos(web, sfolder, LectoresEmb, RoleType.None);
                            this.AsignarPermisos(web, sfolder, LectoresPed, RoleType.Reader);
                            break;
                        case SubCarpetas.Extras:
                            this.AsignarPermisos(web, sfolder, ColaboradoresExpo, RoleType.Contributor);
                            this.AsignarPermisos(web, sfolder, PropExp, RoleType.Contributor);
                            break;
                        default:
                            this.AsignarPermisos(web, sfolder, ColaboradoresExpo, RoleType.Contributor);
                            this.AsignarPermisos(web, sfolder, PropExp, RoleType.Contributor);
                            this.AsignarPermisos(web, sfolder, LectoresPed, RoleType.Reader);
                            break;
                    }
                }
            }
        }


        public List<string> ObtenerNombresSubcarpetas(string url)
        {

            ClientContext clientContext = this.context;
            Web web = clientContext.Web;
            clientContext.Load(web);

            List<string> retornar = new List<string>();

            List listaSubcarpetas = web.Lists.GetByTitle(Listas.Subcarpetas);
            clientContext.Load(listaSubcarpetas);

            CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
            ListItemCollection itemsSubcarpetas = listaSubcarpetas.GetItems(query);

            clientContext.Load(itemsSubcarpetas);
            clientContext.ExecuteQuery();

            if (listaSubcarpetas != null)
            {
                foreach (ListItem item in itemsSubcarpetas)
                {
                    clientContext.Load(item, it => it["LinkTitle"]);
                    clientContext.ExecuteQuery();
                    retornar.Add(item["LinkTitle"].ToString());
                }
            }
                
            return retornar;
        }

        public void AsignarPermisos(Web web, ListItem item, Principal grupo, RoleType roleType)
        {
            ClientContext clientContext = this.context;
            clientContext.Load(item, it => it.HasUniqueRoleAssignments);
            clientContext.ExecuteQuery();

            if (!item.HasUniqueRoleAssignments)
                item.BreakRoleInheritance(false,false);

            clientContext.Load(web);
            clientContext.ExecuteQuery();

            UserCollection users = web.SiteUsers;
            RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(web.Context);
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByType(roleType);
            collRoleDefinitionBinding.Add(roleDefinition);

            item.RoleAssignments.Add(grupo, collRoleDefinitionBinding);

           
            clientContext.Load(item, it => it.RoleAssignments);
            clientContext.Load(users);
            clientContext.ExecuteQuery();

            bool existsysuser = false;
            foreach (RoleAssignment itemfe in item.RoleAssignments)
            {
                clientContext.Load(itemfe, it => it.Member);
                clientContext.Load(itemfe.Member, itm => itm.LoginName);
                clientContext.ExecuteQuery();
                if (itemfe.Member.LoginName.Equals("SHAREPOINT\\system"))
                    existsysuser = true;
            }

            if (existsysuser)
            {
                Principal sysuser = users.GetByLoginName("SHAREPOINT\\system");
                RoleAssignment ra = item.RoleAssignments.GetByPrincipal(sysuser);
                ra.DeleteObject();
            }
        }

        public void AsignarPermisosPorRoleDefinition(Web web, ListItem item, Principal grupo, RoleDefinition roleDefinition)
        {
            ClientContext clientContext = this.context;
            clientContext.Load(item, it => it.HasUniqueRoleAssignments);
            clientContext.ExecuteQuery();

            if (!item.HasUniqueRoleAssignments)
                item.BreakRoleInheritance(false,false);

            UserCollection users = web.SiteUsers;
            RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(web.Context);
            collRoleDefinitionBinding.Add(roleDefinition);

            item.RoleAssignments.Add(grupo,collRoleDefinitionBinding);

            clientContext.Load(item, it => it.RoleAssignments);
            clientContext.Load(users);
            clientContext.ExecuteQuery();

            bool existsysuser = false;
            foreach (RoleAssignment itemfe in item.RoleAssignments)
            {
                clientContext.Load(itemfe, it => it.Member);
                clientContext.Load(itemfe.Member, itm => itm.LoginName);
                clientContext.ExecuteQuery();
                if (itemfe.Member.LoginName.Equals("SHAREPOINT\\system"))
                    existsysuser = true;
            }

            if (existsysuser)
            {
                Principal syuser = users.GetByLoginName("SHAREPOINT\\system");
                RoleAssignment ra = item.RoleAssignments.GetByPrincipal(syuser);
                ra.DeleteObject();
            }
        }

        /* PERMISOS ELEVADOS */
        public ListItem ObtenerClienteParaLookup(string cliente)
        {
            ListItem itemcliente = null;

            using (ClientContext clientContext = this.context)
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                List listaclientes = web.Lists.GetByTitle(Listas.Clientes);

                CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                ListItemCollection itemsClientes = listaclientes.GetItems(query);

                clientContext.Load(itemsClientes);
                clientContext.ExecuteQuery();

                foreach (ListItem item in itemsClientes)
                {
                    if (item[Campos.Titulo].ToString() == cliente)
                    {
                        itemcliente = item;
                        break;
                    }
                }                  
            }
            return itemcliente;
        }

        public string ObtenerUrlLista(string nombre)
        {

            using (ClientContext clientContext = this.context)
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                List lista = web.Lists.GetByTitle(nombre);
                clientContext.Load(lista, milista => milista.DefaultViewUrl);

                clientContext.ExecuteQuery();

                if (lista != null)
                    return lista.DefaultViewUrl;
                else
                    return string.Empty;
            }
        }

        public string ObtenerUrlSitio()
        {       
            context.Load(this.context.Web, web => web.Url);
            context.ExecuteQuery();
            return context.Web.Url;
        }

        /* PERMISOS ELEVADOS */
        public List<string> ObtenerClientes()
        {
            List<string> clients = new List<string>();
            clients.Add(" ");

            using (ClientContext clientContext = this.context)
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                List clientes = web.Lists.GetByTitle(Listas.Clientes);

                CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                ListItemCollection itemsClientes = clientes.GetItems(query);

                clientContext.Load(itemsClientes);
                clientContext.ExecuteQuery();

                if (clientes != null)
                {
                    foreach (ListItem cliente in itemsClientes)
                        if (cliente[Campos.Titulo] != null)
                            clients.Add(cliente[Campos.Titulo].ToString());
                }
            }          
            return clients;
        }

        /*PERMISOS ELEVADOS*/
        public bool ExistePedido(string pedido)
        {
            bool retornar = false;

                using (var clientContext = this.context)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    var pedidos = clientContext.Web.Lists.GetByTitle(Listas.CarpetaRaiz);
                     clientContext.Load(pedidos);
                    FolderCollection carpetas = pedidos.RootFolder.Folders;

                    clientContext.Load(carpetas);
                    clientContext.ExecuteQuery();

                    foreach (Folder carpeta in carpetas)
                    {
                        if (carpeta.Name == pedido)
                            retornar = true;
                    }                   
                }           
            return retornar;
        }

        /*PERMISOS ELEVADOS*/
        public bool UsuarioActualTienePermisos()
        {
            bool tienePermisos = false;

            using (ClientContext clientContext = this.context)
            {
                Web web = clientContext.Web;
                clientContext.Load(web);

                User usuario = web.CurrentUser;
                clientContext.Load(usuario);
                GroupCollection gruposALosQuePertenece = usuario.Groups;
                clientContext.Load(gruposALosQuePertenece);
                clientContext.ExecuteQuery();

                foreach (Group grupo in gruposALosQuePertenece)
                {
                    if (grupo.Title == Grupos.ColaboradoresExportaciones || grupo.Title == Grupos.LectoresEmbarque
                        || grupo.Title == Grupos.LectoresPedidos || grupo.Title == Grupos.PropietariosExportaciones)
                        tienePermisos = true;
                }
            }   
            
            return tienePermisos;
        }

        public string NombreSitio()
        {
            context.Load(this.context.Web, web => web.Title);
            context.ExecuteQuery();
            return this.context.Web.Title;
        }

        //Le pasaria por parametro por ahora el nombre de la lista y el ID del item a cambiar
        public void ItemUpdated(string listName, int ID, string pedido)
        {
            ClientContext clientContext = this.context;
            Web web = clientContext.Web;
            clientContext.Load(web);

            List list = web.Lists.GetByTitle(Listas.CarpetaRaiz);
            clientContext.Load(list, li => li.ContentTypes);
            clientContext.ExecuteQuery();

            if (listName.Equals(Listas.CarpetaRaiz))
            {
                ListItem item = list.GetItemById(ID);
                clientContext.Load(item, it => it.Folder.Folders, it => it.ContentType, it=> it[Campos.Pedido], it => it.Folder);
                clientContext.ExecuteQuery();

                if (item.ContentType.Name.Equals("CT.Raiz"))
                {
                    FolderCollection folders = item.Folder.Folders;
                    clientContext.Load(folders);
                    clientContext.ExecuteQuery();

                    foreach (Folder spfColl in folders)
                    {
                        clientContext.Load(spfColl, col => col.ListItemAllFields, kol => kol.ListItemAllFields[Campos.Pedido]);
                        clientContext.ExecuteQuery();

                        spfColl.ListItemAllFields[Campos.Pedido] = Int64.Parse(pedido);
                        ContentTypeCollection ctc = list.ContentTypes;
                        ContentType contentType_sub = Enumerable.FirstOrDefault(ctc, ct => ct.Name == "CT.Sub");
                        spfColl.ListItemAllFields["ContentTypeId"] = contentType_sub.Id;
                        spfColl.ListItemAllFields.Update();
                    }

                    FolderCollection files = item.Folder.Folders;
                    FileCollection fcol = item.Folder.Files;
                    clientContext.Load(files);
                    clientContext.Load(fcol);
                    clientContext.ExecuteQuery();

                    foreach (File f in fcol)
                    {
                        clientContext.Load(f, myf => myf.ListItemAllFields);
                        clientContext.ExecuteQuery();

                        var listItem = f.ListItemAllFields;
                        listItem[Campos.Pedido] = Int64.Parse(pedido);
                        listItem.Update();
                    }

                    foreach (Folder spfolder in files)
                    {
                        clientContext.Load(spfolder, spf => spf.Files);
                        clientContext.ExecuteQuery();
                        foreach (File spfile in spfolder.Files)
                        {
                           // clientContext.Load(spfolder);
                            clientContext.Load(spfile, file => file.ListItemAllFields);
                            clientContext.ExecuteQuery();

                            var listItem = spfile.ListItemAllFields;
                            listItem[Campos.Pedido] = Int64.Parse(pedido);
                            listItem.Update();
                        }

                    }

                    item[Campos.Nombre] = pedido.ToString();
                    item[Campos.Pedido] = Int64.Parse(pedido);
                    item.Update();
                    string nuevoNumero = pedido.ToString();
                    nuevoNumero = nuevoNumero.Replace(".", "");
                   item["FileLeafRef"] = nuevoNumero;
                   item.Update();
                   clientContext.ExecuteQuery();
                }
            }
        }
    }

}

