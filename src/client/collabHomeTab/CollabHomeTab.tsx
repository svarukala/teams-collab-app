import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { ImageFit, Pivot, PivotItem, Image} from 'office-ui-fabric-react';
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as msal from "@azure/msal-browser";
import jwtDecode from "jwt-decode";
import SPOReusable from "./SPOReusable";
import MSGReusable from "./MSGReusable";

let currentAccount: msal.AccountInfo = null;

const msalConfig = {  
    auth: {  
      clientId: process.env.TAB_APP_ID as string, //'c613e0d1-161d-4ea0-9db4-0f11eeabc2fd',
      authority: "https://login.microsoftonline.com/m365x229910.onmicrosoft.com",
      redirectUri: 'https://sridev.ngrok.io/auth-end'
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
    }  
  };
  
const msalInstance = new msal.PublicClientApplication(msalConfig);

const tokenrequest: msal.SilentRequest = {
    //scopes: ["https://m365x229910.sharepoint.com/AllSites.Read", "https://m365x229910.sharepoint.com/AllSites.Manage"],
    scopes: ['User.Read','Sites.ReadWrite.All', 'Files.ReadWrite.All'],
    account: currentAccount,
    };

const loginRequest = {
    //scopes: ["https://m365x229910.sharepoint.com/AllSites.Read", "https://m365x229910.sharepoint.com/AllSites.Manage"]
    scopes: ['User.Read','Sites.ReadWrite.All', 'Files.ReadWrite.All']
  };

const oboRequest = {
    scopes: ["api://3271e1a1-0da7-476b-b573-e360600674a9/access_as_user"],
    account: currentAccount,
};

/**
 * Implementation of the collab home content page
 */
export const CollabHomeTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [ssoToken, setSsoToken] = useState<string>();
    const [accessToken, setAccessToken] = useState<string>();

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                    setName(decoded!.name);
                    setSsoToken(token);
                    getTokenForOboBroker();
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.TAB_APP_URI as string]
            });
        } else {
            setEntityId("Not in Microsoft Teams");
            getTokenForOboBroker();
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
        }
    }, [context]);

    const getTokenForOboBroker = () => {
        if(accessToken){}
        else{
        if( msalInstance.getAllAccounts().length > 0 ) {
            oboRequest.account = msalInstance.getAllAccounts()[0];
        }
        msalInstance.acquireTokenSilent(oboRequest).then((val) => {  
          let headers = new Headers();  
          let bearer = "Bearer " + val.accessToken;  
          console.info("BEARER TOKEN: "+ val.accessToken);
          console.info("ID TOKEN: "+ val.idToken);
          const decoded: { [key: string]: any; } = jwtDecode(val.idToken) as { [key: string]: any; };
          setName(decoded!.name);
          setSsoToken(val.idToken);
          setAccessToken(val.accessToken);
          //window.location.reload();
          }).catch((errorinternal) => {  
            console.info("Internal error: "+ errorinternal); 
            msalInstance.loginPopup(oboRequest).then((resp) => {
                console.info("BEARER TOKEN FROM POPUP: "+ resp.accessToken);
                console.info("ID TOKEN FROM POPUP: "+ resp.idToken);
                const decoded: { [key: string]: any; } = jwtDecode(resp.idToken) as { [key: string]: any; };
                setName(decoded!.name);
                setSsoToken(resp.idToken);
                setAccessToken(resp.accessToken);
            }).catch(e => {
              console.info(e);
            });
          });
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={`Hello ${name}`} />
                        </div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                        {
                            accessToken &&
                            <Pivot aria-label="Basic Pivot Example">
                                <PivotItem headerText="SPO REST API">
                                    <SPOReusable idToken={accessToken} />
                                </PivotItem>
                                <PivotItem headerText="MS Graph REST API">
                                    <MSGReusable idToken={accessToken} />
                                </PivotItem>                                
                            </Pivot>
                        }
                        </div>
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Contoso" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
