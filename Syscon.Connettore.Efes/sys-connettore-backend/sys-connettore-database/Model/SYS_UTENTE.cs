//------------------------------------------------------------------------------
// <auto-generated>
//    Codice generato da un modello.
//
//    Le modifiche manuali a questo file potrebbero causare un comportamento imprevisto dell'applicazione.
//    Se il codice viene rigenerato, le modifiche manuali al file verranno sovrascritte.
// </auto-generated>
//------------------------------------------------------------------------------

namespace sys_connettore_database.Model
{
    using System;
    using System.Collections.Generic;
    
    public partial class SYS_UTENTE
    {
        public int SYS_ID { get; set; }
        public string SYS_USERNAME { get; set; }
        public string SYS_PASSWORD { get; set; }
        public string SYS_DESCRIZIONE { get; set; }
        public string SYS_EMAIL { get; set; }
        public int SYS_ADMIN { get; set; }
        public int SYS_FLGATTIVO { get; set; }
    }
}