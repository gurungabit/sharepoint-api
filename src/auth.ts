import { ConfidentialClientApplication } from '@azure/msal-node';
import type { Configuration } from '@azure/msal-node';
import { config } from './config.js';

let accessTokenCache: { token: string; expiresOn: number } | null = null;

export const getAccessToken = async (): Promise<string> => {
    // Return cached token if it's still valid (with a 5-minute buffer)
    if (accessTokenCache && accessTokenCache.expiresOn > Date.now() + 300000) {
        return accessTokenCache.token;
    }

    if (!config.tenantId || !config.clientId || !config.clientSecret) {
        throw new Error("Missing Azure AD credentials in environment variables.");
    }

    const msalConfig: Configuration = {
        auth: {
            clientId: config.clientId,
            authority: `https://login.microsoftonline.com/${config.tenantId}`,
            clientSecret: config.clientSecret,
        }
    };

    const cca = new ConfidentialClientApplication(msalConfig);

    const clientCredentialRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
    };

    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    
    if (!response || !response.accessToken || !response.expiresOn) {
        throw new Error("Failed to acquire access token.");
    }
    
    // Cache the token
    accessTokenCache = {
        token: response.accessToken,
        expiresOn: response.expiresOn.getTime()
    };
    
    return response.accessToken;
};