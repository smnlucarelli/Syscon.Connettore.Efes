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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    using System.Data.Objects;
    using System.Data.Objects.DataClasses;
    using System.Linq;
    
    public partial class AlyEfesEntities : DbContext
    {
        public AlyEfesEntities()
            : base("name=AlyEfesEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public DbSet<GFTA_CAUSALE> GFTA_CAUSALE { get; set; }
        public DbSet<GFTA_CAUSALEPREST> GFTA_CAUSALEPREST { get; set; }
        public DbSet<GFTA_COMUNE> GFTA_COMUNE { get; set; }
        public DbSet<GFTA_IVA> GFTA_IVA { get; set; }
        public DbSet<GFTA_PAGAMENTO> GFTA_PAGAMENTO { get; set; }
        public DbSet<GFTA_PDC> GFTA_PDC { get; set; }
        public DbSet<GFTA_STATOESTERO> GFTA_STATOESTERO { get; set; }
        public DbSet<GOAR_CLIFOR> GOAR_CLIFOR { get; set; }
        public DbSet<GOCG_PRIMANOTA> GOCG_PRIMANOTA { get; set; }
        public DbSet<GOCM_DOCUMENTO> GOCM_DOCUMENTO { get; set; }
        public DbSet<GORP_RITENUTEP> GORP_RITENUTEP { get; set; }
        public DbSet<VW_GFAR_CLIFOR> VW_GFAR_CLIFOR { get; set; }
        public DbSet<SYS_MENU> SYS_MENU { get; set; }
        public DbSet<SYS_PARAMETRI> SYS_PARAMETRI { get; set; }
        public DbSet<SYS_UTENTE> SYS_UTENTE { get; set; }
    
        public virtual int SP_CreaRecord()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_CreaRecord");
        }
    
        public virtual int SP_GFTA_CAUSALE()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_CAUSALE");
        }
    
        public virtual int SP_GFTA_CAUSALEPREST()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_CAUSALEPREST");
        }
    
        public virtual int SP_GFTA_COMUNE()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_COMUNE");
        }
    
        public virtual int SP_GFTA_IVA()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_IVA");
        }
    
        public virtual int SP_GFTA_PAGAMENTO()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_PAGAMENTO");
        }
    
        public virtual int SP_GFTA_PDC()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_PDC");
        }
    
        public virtual int SP_GFTA_STATOESTERO()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("SP_GFTA_STATOESTERO");
        }
    
        public virtual int zSP_CANCELLA_TABELLE_GF()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("zSP_CANCELLA_TABELLE_GF");
        }
    
        public virtual int zSP_GOCG_CliFor()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("zSP_GOCG_CliFor");
        }
    }
}
