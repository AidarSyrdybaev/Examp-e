using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp3.Entities
{
    public class Client: IEntity
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime BirthDate { get; set; }
        public string PhoneNumber { get; set; }
        public string Address { get; set; }
        public string SocialNumber { get; set; }
    }
}
