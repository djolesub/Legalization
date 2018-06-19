namespace Resenja_EF_Konzistentno_DIV
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("RR")]
    public partial class RR
    {
        public int? PO_id { get; set; }

        [StringLength(255)]
        public string PO_naziv { get; set; }

        public int? KO_id { get; set; }

        [StringLength(255)]
        public string KO_naziv { get; set; }

        public string brojParcele { get; set; }

        public string podbrojParcele { get; set; }

        public string punBrojParcele { get; set; }

        public string NepId { get; set; }

        [StringLength(50)]
        public string vrstaRR { get; set; }

        public byte? izvezen { get; set; }

        [StringLength(255)]
        public string dokument { get; set; }

        public string dokumentId { get; set; }

        public int? redniBroj { get; set; }

        [StringLength(255)]
        public string originalNaziv { get; set; }

        [StringLength(255)]
        public string napomena { get; set; }

        public DateTime? datumKreiranja { get; set; }

        public bool FajlDaNe { get; set; }

        public int ID { get; set; }

        //Defining Methods 
        public void SpojBrojIPodbrojParcele()
        {
            if (this.podbrojParcele == "000")
            {
                this.punBrojParcele = this.brojParcele.TrimStart(new char[] { '0' });

            }
            else
            {
                this.punBrojParcele = String.Format(@"{0}/{1}", this.brojParcele.TrimStart(new char[] { '0' }), this.podbrojParcele.TrimStart(new char[] { '0' }));
                //Console.WriteLine(this.punBrojParcele);
            }
        }

        public void DopuniBrojParcele()
        {
            int broj_parcele_duzina = this.brojParcele.Length;

            switch (broj_parcele_duzina)
            {
                case 1:
                    this.brojParcele = String.Format(@"0000{0}", this.brojParcele);
                    break;
                case 2:
                    this.brojParcele = String.Format(@"000{0}", this.brojParcele);
                    break;
                case 3:
                    this.brojParcele = String.Format(@"00{0}", this.brojParcele);
                    break;
                case 4:
                    this.brojParcele = String.Format(@"0{0}", this.brojParcele);
                    break;
                case 5:
                    this.brojParcele = String.Format(@"{0}", this.brojParcele);
                    break;

                default:
                    Console.WriteLine("Default Case Dopuni Broj Parcele");
                    break;
            }
        }
        public void DopuniPodrojParcele()
        {
            int broj_parcele_duzina = this.podbrojParcele.Length;
            switch (broj_parcele_duzina)
            {
                case 1:
                    this.podbrojParcele = String.Format(@"00{0}", this.podbrojParcele);
                    break;
                case 2:
                    this.podbrojParcele = String.Format(@"0{0}", this.podbrojParcele);
                    break;
                case 3:
                    this.podbrojParcele = String.Format(@"{0}", this.podbrojParcele);
                    break;
                default:
                    Console.WriteLine("Default Case Dopuni PodBroj Parcele");
                    break;
            }
        }

        public void KreirajNepID()
        {
            this.NepId = String.Format(@"{0}{1}{2}", this.KO_id, this.brojParcele, this.podbrojParcele);
        }

        public void KreirajDokumentID()
        {
            this.dokumentId = this.dokument.Split(new char[] { '.' })[0];
        }
    }
}
