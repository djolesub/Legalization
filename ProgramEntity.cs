using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Entity;

namespace Resenja_EF_Konzistentno_DIV
{
    class Program
    {
        static void Main(string[] args)
        {
            using (Resenja context = new Resenja())
            {
                context.Database.CreateIfNotExists();
                Console.WriteLine("SQL Server Central Database..");
                Console.WriteLine("number OF Records is: {0}", context.RR.Count());
                DbSet<RR> rr = context.RR;
                foreach (RR r in rr)
                {

                    r.DopuniBrojParcele();
                    r.DopuniPodrojParcele();
                    r.SpojBrojIPodbrojParcele();
                    r.KreirajDokumentID();
                    r.KreirajNepID();
                }
                context.SaveChanges();
                Console.WriteLine("Finishing succesfuly");

            }
        }
    }
}
