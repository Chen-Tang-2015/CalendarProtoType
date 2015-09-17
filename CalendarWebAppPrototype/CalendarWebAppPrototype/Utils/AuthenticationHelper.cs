using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.OutlookServices;
using System;
using System.Security.Claims;
using System.Threading.Tasks;
using CalendarWebAppPrototype.Models;
using Microsoft.Azure.ActiveDirectory.GraphClient;

namespace CalendarWebAppPrototype.Utils
{
    internal class AuthenticationHelper
    {
        private static ActiveDirectoryClient _graphClient = null;
        private const string AadServiceResourceId = "https://graph.windows.net/";
        private static readonly string CommonAuthority = "https://login.microsoftonline.com/Common";
        private static string LastAuthority { get; set; }
        private static string TenantId { get; set; }
        private static AuthenticationContext _authenticationContext { get; set; }
        #region Get Service Client

        public static async Task<GraphService> GetGraphServiceAsync()
        {
            var serviceRoot = GetGraphServiceRoot();

            var accessToken = GetAccessTokenAsync(SettingsHelper.GraphResourceId);
            // AdalException thrown by GetAccessTokenAsync is swallowed 
            // by GraphService so we need to wait here.
            await accessToken;
            return new GraphService(serviceRoot, () => accessToken);
        }

        public static ActiveDirectoryClient GetGraphClient()
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_graphClient != null)
            {
                return _graphClient;
            }

            // Active Directory service endpoints
            Uri aadServiceEndpointUri = new Uri(AadServiceResourceId);

            try
            {
                //First, look for the authority used during the last authentication.
                //If that value is not populated, use CommonAuthority.
                string authority = String.IsNullOrEmpty(LastAuthority) ? CommonAuthority : LastAuthority;

                // Create an AuthenticationContext using this authority.
                _authenticationContext = new AuthenticationContext(authority);

                var token = GetAccessToken(AadServiceResourceId);

                // Check the token
                if (string.IsNullOrEmpty(token))
                {
                    // User cancelled sign-in
                    throw new Exception("Sign-in cancelled");  // assuming we don't want to continue
                }
                else
                {
                    // Create our ActiveDirectory client.
                    _graphClient = new ActiveDirectoryClient(
                        new Uri(aadServiceEndpointUri, TenantId),
                        async () => await GetAccessTokenAsync(AadServiceResourceId)
                        );

                    return _graphClient;
                }
            }
            catch (Exception)
            {
                _authenticationContext.TokenCache.Clear();
                throw;
            }
        }

        public static async Task<OutlookServicesClient> GetOutlookServiceAsync()
        {
            var serviceRoot = new Uri(SettingsHelper.OutlookResourceUri);
            var accessToken = GetAccessTokenAsync(SettingsHelper.OutlookResourceId);
            // AdalException thrown by GetAccessTokenAsync ise swallowed 
            // by EntityContainer so we need to wait here.
            await accessToken;
            return new OutlookServicesClient(serviceRoot, () => accessToken);
        }

        #endregion

        #region Get Access Token

        public static string GetAccessToken(string resource)
        {
            var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey);

            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);
            var context = GetAuthenticationContext();
            try
            {
                var token = context.AcquireTokenSilent(resource, clientCredential, userIdentifier);
                return token.AccessToken;
            }
            catch (AdalException exception)
            {
                //Partially handle token acquisition failure here and bubble it up to the controller
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    context.TokenCache.Clear();
                    throw exception;
                }
                return null;
            }
        }

        public static async Task<string> GetAccessTokenAsync(string resourceId)
        {
            var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey);

            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);
            var context = GetAuthenticationContext();
            try
            {
                var token = await context.AcquireTokenSilentAsync(resourceId, clientCredential, userIdentifier);
                return token.AccessToken;
            }
            catch (AdalException exception)
            {
                //Partially handle token acquisition failure here and bubble it up to the controller
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    context.TokenCache.Clear();
                    throw exception;
                }
                return null;
            }

        }

        public static async Task<string> GetGraphAccessTokenAsync()
        {
            return await GetAccessTokenAsync(SettingsHelper.GraphResourceId);
        }

        public static async Task<string> GetOutlookAccessTokenAsync()
        {
            return await GetAccessTokenAsync(SettingsHelper.OutlookResourceId);
        }

        #endregion

        #region Private methods

        private static Uri GetGraphServiceRoot()
        {
            var servicePointUri = new Uri(SettingsHelper.GraphResourceUri);
            var tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            return new Uri(servicePointUri, tenantId);
        }

        private static AuthenticationContext GetAuthenticationContext()
        {
            var tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            var authority = string.Format("{0}/{1}", SettingsHelper.AuthorizationUri, tenantId);

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var tokenCache = new ADALTokenCache(signInUserId);

            return new AuthenticationContext(authority, tokenCache);
        }

        #endregion
    }
}