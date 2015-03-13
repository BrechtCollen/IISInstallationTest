using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IISInstallationTest.models
{
    public class ImpersonatedUser
    {
        public string Name { get; set; }
        public string Password { get; set; }

        public void GetName()
        {
            Console.WriteLine(@"Enter the username in the form DOMAIN\username");
            Name = Console.ReadLine();
            //Name = @"INTRANET\bco";
        }

        public void GetPassword()
        {
            Console.WriteLine("Enter the password");
            Password = Console.ReadLine();
            //Password = "S456789-";
        }
    }
}
