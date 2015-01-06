using System;
using System.Linq;
using NetFwTypeLib;

namespace Kesoft.Windows.NetFirewallHelper
{
    /// <summary>
    /// 网络防火墙管理器。
    /// </summary>
    public static class NetFirewall
    {
        /// <summary>
        /// 端口协议。
        /// </summary>
        public enum Protocol
        {
            TCP,
            UDP
        }

        /// <summary>
        /// 添加防火墙例外端口。
        /// </summary>
        /// <param name="name">名称。</param>
        /// <param name="port">端口。</param>
        /// <param name="protocol">协议。</param>
        public static void AddPort(string name, int port, Protocol protocol)
        {
            if (!PortIsAuthorized(port, protocol))
            {
                var manager = GetNetFirewallNManager();
                var op = (INetFwOpenPort) Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwOpenPort"));
                op.Name = name;
                op.Port = port;
                op.Protocol = protocol == Protocol.TCP ? NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP : NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_UDP;
                op.Scope = NET_FW_SCOPE_.NET_FW_SCOPE_ALL;
                op.Enabled = true;
                manager.LocalPolicy.CurrentProfile.GloballyOpenPorts.Add(op);
            }
        }

        /// <summary>
        /// 删除防火墙例外端口。
        /// </summary>
        /// <param name="port">端口。</param>
        /// <param name="protocol">协议（TCP、UDP）。</param>
        public static void RemovePort(int port, Protocol protocol)
        {
            var manager = GetNetFirewallNManager();
            manager.LocalPolicy.CurrentProfile.GloballyOpenPorts.Remove(port, protocol == Protocol.TCP ? NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP : NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_UDP);
        }

        /// <summary>
        /// 添加应用程序到防火墙例外。
        /// </summary>
        /// <param name="executablePath">应用程序的绝对全路径可执行文件名称。</param>
        public static void AddApplication(string executablePath)
        {
            if (!ApplicationIsAuthorized(executablePath))
            {
                var manager = GetNetFirewallNManager();
                var app = (INetFwAuthorizedApplication) Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwAuthorizedApplication"));
                app.Name = executablePath.ToLower();
                app.ProcessImageFileName = executablePath.ToLower();
                app.Enabled = true;
                manager.LocalPolicy.CurrentProfile.AuthorizedApplications.Add(app);
            }
        }

        /// <summary>
        /// 删除防火墙例外应用程序。
        /// </summary>
        /// <param name="executablePath">应用程序的绝对全路径可执行文件名称。</param>
        public static void RemoveApplication(string executablePath)
        {
            var manager = GetNetFirewallNManager();
            manager.LocalPolicy.CurrentProfile.AuthorizedApplications.Remove(executablePath);
        }

        /// <summary>
        /// 端口是否已授权。
        /// </summary>
        /// <param name="port">端口。</param>
        /// <param name="protocol">协议。</param>
        /// <returns></returns>
        public static bool PortIsAuthorized(int port, Protocol protocol)
        {
            var manager = GetNetFirewallNManager();
            var p = protocol == Protocol.TCP ? NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP : NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_UDP;
            return manager.LocalPolicy.CurrentProfile.GloballyOpenPorts.Cast<INetFwOpenPort>().Any(a => a.Port == port && a.Protocol == p);
        }

        /// <summary>
        /// 应用程序是否已经授权。
        /// </summary>
        /// <param name="executablePath">应用程序的绝对全路径可执行文件名称。</param>
        /// <returns></returns>
        public static bool ApplicationIsAuthorized(string executablePath)
        {
            return GetNetFirewallNManager()
                .LocalPolicy.CurrentProfile
                .AuthorizedApplications
                .Cast<INetFwAuthorizedApplication>()
                .Any(a => a.ProcessImageFileName.ToLower() == executablePath.ToLower());
        }

        /// <summary>
        /// 获取防火墙管理对象。
        /// </summary>
        /// <returns>网络防火墙管理对象。</returns>
        private static INetFwMgr GetNetFirewallNManager()
        {
            return (INetFwMgr) Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwMgr"));
        }
    }
}