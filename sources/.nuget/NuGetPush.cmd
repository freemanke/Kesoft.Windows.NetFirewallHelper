nuget update -self
nuget SetApiKey 11b952d1-5fa1-457c-adad-ec28a8c50bb4
nuget pack ..\Kesoft.Windows.NetFirewallHelper\Kesoft.Windows.NetFirewallHelper.csproj
nuget push Kesoft.Windows.NetFirewallHelper.*.nupkg