using sys_connettore_database.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace sys_connettore_webapi.Services
{
    
    public interface ISYLoginService
    {
        public class SYS_UTENTI_EXT
        {
            public int SYS_ID { get; set; }
            public string SYS_USERNAME { get; set; }
            public string SYS_PASSWORD { get; set; }
            public string SYS_DESCRIZIONE { get; set; }
            public string SYS_EMAIL { get; set; }
            public int SYS_ADMIN { get; set; }
            public int SYS_FLGATTIVO { get; set; }
        }


        Tuple<bool, SYS_UTENTI_EXT> UserLogin(string username, string password);


        public class SYLoginService : ISYLoginService 
            {
            public Tuple<bool, SYS_UTENTI_EXT> UserLogin(string username, string password)
            {
                bool login = false;
                SYS_UTENTI_EXT user = new SYS_UTENTI_EXT();

                using (var ctx = new ServizioEntities())
                {
                    user = ctx.SYS_UTENTI.Where(x => x.SYS_USERNAME.Trim() == username)
                                         .Select(x => new SYS_UTENTI_EXT
                                         {
                                             SYS_ID = x.SYS_ID,
                                             SYS_USERNAME = x.SYS_USERNAME,
                                             SYS_DESCRIZIONE = x.SYS_DESCRIZIONE,
                                             SYS_PASSWORD = x.SYS_PASSWORD,
                                             SYS_EMAIL = x.SYS_EMAIL,
                                             SYS_ADMIN = x.SYS_ADMIN,
                                             SYS_FLGATTIVO = x.SYS_FLGATTIVO
                                         })
                                         .FirstOrDefault();
                }

                if (user == null)
                {
                    login = false;
                    SYS_UTENTI_EXT anyone = new SYS_UTENTI_EXT();

                    return Tuple.Create(login, anyone);

                }
                else
                {
                    if (user.SYS_PASSWORD.Trim() == password)
                    {
                        login = true;
                        return Tuple.Create(login, user);
                    }
                    else if (user.SYS_PASSWORD.Trim() != password)
                    {
                        login = false;
                        SYS_UTENTI_EXT anyone = new SYS_UTENTI_EXT();

                        return Tuple.Create(login, anyone);

                    }
                    else if (user.SYS_PASSWORD.Trim() == null || user.SYS_PASSWORD.Trim() == "")
                    {
                        login = false;
                        SYS_UTENTI_EXT anyone = new SYS_UTENTI_EXT();

                        return Tuple.Create(login, anyone);
                    }

                }

                return Tuple.Create(login, user);
            }

        }

    }

}
