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
    
    public partial class GOAR_CLIFOR
    {
        public long ID { get; set; }
        public Nullable<System.Guid> Id_Elaborazione_1 { get; set; }
        public decimal Ditta { get; set; }
        public decimal Tipocf { get; set; }
        public Nullable<decimal> Key_1 { get; set; }
        public decimal Stato { get; set; }
        public decimal Esito { get; set; }
        public string PartIva { get; set; }
        public string CodFiscale { get; set; }
        public decimal FlgPrsFis { get; set; }
        public string Cognome { get; set; }
        public string Nome { get; set; }
        public decimal Sesso { get; set; }
        public Nullable<System.DateTime> DataNascita { get; set; }
        public string ComNascita { get; set; }
        public string ProvNascita { get; set; }
        public string CodiceComuneNasc { get; set; }
        public Nullable<decimal> NazioneIsoNasc { get; set; }
        public string CodPag { get; set; }
        public string Tel1Num { get; set; }
        public string Tel2Num { get; set; }
        public string FaxNum { get; set; }
        public string CellNum { get; set; }
        public string IndirizzoEmail { get; set; }
        public string EmailPec { get; set; }
        public string RagioneSociale { get; set; }
        public string CodiceComune { get; set; }
        public string Indirizzo { get; set; }
        public string Citta { get; set; }
        public string Cap { get; set; }
        public string Prov { get; set; }
        public Nullable<decimal> NazioneIso { get; set; }
        public string RagSocFisc { get; set; }
        public string CodiceComuneFisc { get; set; }
        public string IndFisc { get; set; }
        public string CittaFisc { get; set; }
        public string CapFisc { get; set; }
        public string ProvFisc { get; set; }
        public Nullable<decimal> NazioneIsoFisc { get; set; }
        public string PartIvaEst { get; set; }
        public string Aliva { get; set; }
        public decimal IndSoggRit { get; set; }
        public string CodiceCauPrest { get; set; }
        public Nullable<byte> IndTipoInstradamento { get; set; }
        public string CodDestSdi { get; set; }
        public Nullable<decimal> IndStato { get; set; }
    }
}
