using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using NUnit.Framework;

namespace Kesoft.Windows.NetFirewallHelper.Tests
{
    [TestFixture]
    public class NetFirewallTest
    {
        [Test]
        public void AddRemovePort()
        {
            const string name = "tests";
            const int port = 4556;
            foreach (var p in new[] {NetFirewall.Protocol.TCP, NetFirewall.Protocol.UDP})
            {
                var authorized = NetFirewall.PortIsAuthorized(port, p);
                if (authorized) NetFirewall.RemovePort(port, p);

                NetFirewall.AddPort(name, port, p);
                authorized = NetFirewall.PortIsAuthorized(port, p);
                Assert.IsTrue(authorized);
                NetFirewall.RemovePort(port, p);
                authorized = NetFirewall.PortIsAuthorized(port, p);
                Assert.IsFalse(authorized);
            }
        }

        [Test]
        public void AddRemoveApplication()
        {
            var authorized = NetFirewall.ApplicationIsAuthorized(Application.ExecutablePath);
            if (authorized) NetFirewall.RemoveApplication(Application.ExecutablePath);

            NetFirewall.AddApplication(Application.ExecutablePath);
            authorized = NetFirewall.ApplicationIsAuthorized(Application.ExecutablePath);
            Assert.IsTrue(authorized);
            NetFirewall.RemoveApplication(Application.ExecutablePath);
            authorized = NetFirewall.ApplicationIsAuthorized(Application.ExecutablePath);
            Assert.IsFalse(authorized);
        }
    }
}