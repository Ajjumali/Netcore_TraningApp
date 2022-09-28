using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TrainingApp.Models;

namespace TrainingApp
{
    public class MyConnection: DbContext
    {
        public MyConnection()
        {

        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=172.16.41.226;Database=traningdb;User ID=CMS; password=jgc#2018$Dev");
        }

        public DbSet<PCWPS> PCWBS { get; set; }
        public DbSet<FWBS> Fwbs { get; set; }
        public DbSet<MileStone> MileStone { get; set; }

        
    }
}
