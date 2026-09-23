// Updated original code from
// Added support for custom file location
// https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-net-token-cache-serialization#simple-token-cache-serialization-msal-only

using System;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Identity.Client;

public static class TokenCacheHelperEx
{
    public static void EnableSerialization(ITokenCache tokenCache, String fileName = @"%LOCALAPPDATA%\GraphPowerShellManager\MSALToken.bin")
    {
        tokenCache.SetBeforeAccess(BeforeAccessNotification);
        tokenCache.SetAfterAccess(AfterAccessNotification);

        CacheFilePath = Environment.ExpandEnvironmentVariables(fileName);
    }

    /// <summary>
    /// Path to the token cache
    /// </summary>

    public static string CacheFilePath { get; private set;}

    private static readonly object FileLock = new object();

    private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
    {
        lock (FileLock)
        {
            args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath)
                    ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath),
                    null,
                    DataProtectionScope.CurrentUser)
                    : null);
        }
    }

    private static void AfterAccessNotification(TokenCacheNotificationArgs args)
    {
        // if the access operation resulted in a cache update
        if (args.HasStateChanged)
        {
            lock (FileLock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CacheFilePath));
                // reflect changes in the persistent store
                File.WriteAllBytes(CacheFilePath,
                                    ProtectedData.Protect(args.TokenCache.SerializeMsalV3(),
                                                            null,
                                                            DataProtectionScope.CurrentUser)
                                    );
            }
        }
    }
}