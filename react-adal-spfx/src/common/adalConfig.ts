import { AuthenticationContext, adalFetch, AdalConfig } from 'react-adal';
// App Registration ID
const appId = '477e380a-be19-462c-a3ca-58e9f4b993c7';
export const adalConfig: AdalConfig = {
    cacheLocation: 'localStorage',
    clientId: appId,
    endpoints: {
        "https://graph.microsoft.com": "https://graph.microsoft.com",
        "TestAPI": "https://campusfn.azurewebsites.net"
    },
    postLogoutRedirectUri: window.location.origin,
    tenant: 'anibapatdev.onmicrosoft.com'
};
export const authContext = new AuthenticationContext(adalConfig);

export const adalApiFetch = (endpoint:string , url: string, options: any = {}) => {
    const headers = {
        "accept": "application/json;odata=verbose"
    };
    return adalFetch(authContext, endpoint, fetch, url, headers);
};