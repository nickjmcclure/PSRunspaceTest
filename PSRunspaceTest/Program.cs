using Microsoft.Identity.Client;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Security.Cryptography.X509Certificates;


string clientId = "";
string tenantId = "";
string tenantDomain = "";
string powerShellEndpoint = "https://outlook.office365.com/powershell-liveid/";

string pathToCert = @"";
string certPassword = "";


//Get the Access Token
X509Certificate2 cert = new(pathToCert, certPassword, X509KeyStorageFlags.Exportable);

IConfidentialClientApplication thisApp = ConfidentialClientApplicationBuilder.Create(clientId)
    .WithCertificate(cert)
    .WithAuthority($"https://login.windows.net/{tenantDomain}/")
    .Build();

AuthenticationResult accessToken = thisApp.AcquireTokenForClient(new[] { $"https://outlook.office365.com/.default" }).ExecuteAsync().Result;


//Create the PSCredential
string auth = $"Bearer {accessToken.AccessToken}";
SecureString password = GetSecureString(auth);

PSCredential psCredential = new($"OAuthUser@{tenantId}", password);


//Create the WSManConnectionInfo
string email = $"SystemMailbox{{bb558c35-97f1-4cb9-8ff7-d53741dc928c}}@{tenantDomain}";
string urlEncodedEmail = Uri.EscapeDataString(email);

Uri psURI = new($"{powerShellEndpoint}?BasicAuthToOAuthConversion=true&email={urlEncodedEmail}");

WSManConnectionInfo connectionInfo = new(psURI, "Microsoft.Exchange", psCredential)
{
    AuthenticationMechanism = AuthenticationMechanism.Basic,
};

//Create the Runspace and open the Runspace
using Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
runspace.Open();

using PowerShell ps = PowerShell.Create();


static SecureString GetSecureString(string input)
{
    SecureString password = new();
    foreach (char c in input)
    {
        password.AppendChar(c);
    }
    return password;
}