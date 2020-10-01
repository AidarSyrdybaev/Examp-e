using ConsoleApp3.Entities;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp3.Context
{
    public class ApplicationDbContext: DbContext
    {   
        public DbSet<Client> Clients { get; set; }

        public ApplicationDbContext(DbContextOptions options) : base(options)
        { 
        
        }
    }
}
