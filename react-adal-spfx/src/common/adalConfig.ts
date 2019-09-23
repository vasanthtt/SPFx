import { AuthenticationContext, adalFetch, AdalConfig } from 'react-adal';
// App Registration ID
const appId = '{guid}';
export const adalConfig: AdalConfig = {
    cacheLocation: 'localStorage',
    clientId: appId,
    endpoints: {
        "https://graph.microsoft.com": "https://graph.microsoft.com",
        "TestAPI": "https://fn.azurewebsites.net"
    },
    postLogoutRedirectUri: window.location.origin,
    tenant: '[tenant].onmicrosoft.com'
};
export const authContext = new AuthenticationContext(adalConfig);

export const adalApiFetch = (endpoint:string , url: string, options: any = {}) => {
    const headers = {
        "accept": "application/json;odata=verbose"
    };
    return adalFetch(authContext, endpoint, fetch, url, headers);
};
