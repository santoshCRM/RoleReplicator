using Microsoft.Xrm.Sdk;
using System;

namespace RoleReplicatorControl
{
    internal class SecurityRole
    {
        public string Name { get; set; }
        public EntityReference BusinessUnit { get; set; }

        public Guid RoleId { get; set; }



    }
}
