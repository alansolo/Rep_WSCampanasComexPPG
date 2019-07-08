using EntidadesCampanasPPG.Modelo;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatosCampanasPPG.Catalogo;
using System.Data;
using UtilidadesCampanasPPG;

namespace NegocioCampanasPPG.Ldap
{
    public class LdapNEG
    {
        public UsuarioLdap BuscarUsuario(string Email)
        {
            UsuarioLdap usuarioLdap = new UsuarioLdap();
            string serverLdap = string.Empty;
            string directorioLdap = string.Empty;

            try
            {
                serverLdap = ConfigurationManager.AppSettings["serverLdap"];
                directorioLdap = ConfigurationManager.AppSettings["directorioLdap"];

                using (var context = new PrincipalContext(ContextType.Domain, serverLdap, directorioLdap))
                {
                    UserPrincipal userPrincipal = new UserPrincipal(context);
                    userPrincipal.EmailAddress = Email;

                    using (var searcher = new PrincipalSearcher(userPrincipal))
                    {
                        Principal principal = searcher.FindOne();

                        usuarioLdap.Nombre = principal.Name;
                        usuarioLdap.PPGID = principal.SamAccountName;
                        usuarioLdap.Email = Email;
                    }
                }
            }
            catch (Exception ex)
            {
                usuarioLdap = new UsuarioLdap();

                ArchivoLog.EscribirLog(null, "ERROR: Service - BuscarUsuario, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return usuarioLdap;
        }
        public List<UsuarioLdap> GetUsuarioLdap(List<Parametro> ListParametro)
        {
            List<UsuarioLdap> ListUsurioLdapTemp = new List<UsuarioLdap>();
            List<UsuarioLdap> ListUsurioLdap = new List<UsuarioLdap>();
            string serverLdap = string.Empty;
            string directorioLdap = string.Empty;

            LdapDAT ldapDAT = new LdapDAT();
            DataTable dtDirectorioLdap = new DataTable();
            List<DirectorioActivo> ListDirectorioActivo = new List<DirectorioActivo>();

            ParametroDAT parametroDAT = new ParametroDAT();
            DataTable dtParametro = new DataTable();
            Parametro parametro = new Parametro();

            try
            {
                dtParametro = parametroDAT.GetParametro(0, null);

                ListParametro = dtParametro.AsEnumerable()
                                .Select(n => new Parametro
                                {
                                    Id = n.Field<int?>("Id").GetValueOrDefault(),
                                    Nombre = n.Field<string>("Nombre"),
                                    Valor = n.Field<string>("Valor")
                                }).ToList();

                parametro = ListParametro.Where(n => n.Nombre.ToUpper() == ConfigurationManager.AppSettings["ServerLdap"].ToUpper()).FirstOrDefault();

                dtDirectorioLdap = ldapDAT.GetDirectorioLdap(0, null);

                ListDirectorioActivo = dtDirectorioLdap.AsEnumerable()
                                        .Select(n => new DirectorioActivo
                                        {
                                            Id = n.Field<int?>("Id").GetValueOrDefault(),
                                            Clave = n.Field<string>("Clave"),
                                            Descripcion = n.Field<string>("Descripcion"),
                                            Ldap = n.Field<string>("Ldap")
                                        }).ToList();

                //DIRECTORIO LDAP
                serverLdap = parametro.Valor;

                foreach (DirectorioActivo directorio in ListDirectorioActivo)
                {
                    //DIRECTORIO
                    directorioLdap = directorio.Ldap;

                    using (var context = new PrincipalContext(ContextType.Domain, serverLdap, directorioLdap))
                    {
                        UserPrincipal userPrincipal = new UserPrincipal(context);

                        using (var searcher = new PrincipalSearcher(userPrincipal))
                        {
                            var ListUserPrincipal = searcher.FindAll().Cast<UserPrincipal>().ToList();

                            ListUsurioLdapTemp = ListUserPrincipal.Select(row => new UsuarioLdap
                            {
                                Email = row.EmailAddress != null ? row.EmailAddress.ToLower() : string.Empty,
                                Nombre = row.Name,
                                PPGID = row.SamAccountName
                            }).ToList();

                            ListUsurioLdap.AddRange(ListUsurioLdapTemp);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ArchivoLog.EscribirLog(null, "ERROR: Service - GetUsuarioLdap, Source:" + ex.Source + ", Message:" + ex.Message);

                throw ex;
            }

            return ListUsurioLdap;
        }
    }
}
