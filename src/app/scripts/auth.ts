import * as Msal from "msal";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * Implementation of the teams app1 Auth page
 */
export class Auth {
  private token: string = "";
  private user: Msal.Account;

  /**
   * Constructor for Tab that initializes the Microsoft Teams script
   */
  constructor() {
    microsoftTeams.initialize();
  }

  public performAuthV2(teamsFlow: boolean = true) {
    console.log("Authv2", teamsFlow);
    // Setup auth parameters for MSAL
    const graphAPIScopes: string[] = ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/User.ReadBasic.All", "https://graph.microsoft.com/Group.Read.All"];
    const msalConfig: Msal.Configuration = {
      auth: {
        clientId: "e743c151-a549-4181-b3e9-e84052c9174c",
        authority: "https://login.microsoftonline.com/common" // todo replace with teams context tid
      }
    };

    const userAgentApplication = new Msal.UserAgentApplication(msalConfig);
    userAgentApplication.handleRedirectCallback(() => { const notUsed = ""; });

    if (userAgentApplication.isCallback(window.location.hash)) {

      const user = userAgentApplication.getAccount();
      console.log("Starting callback", user);
      if (user) {
        this.getToken(userAgentApplication, graphAPIScopes, teamsFlow);
      }
    } else {
      this.user = userAgentApplication.getAccount();
      console.log("Starting !callback", this.user);
      if (!this.user) {
        // If user is not signed in, then prompt user to sign in via loginRedirect.
        // This will redirect user to the Azure Active Directory v2 Endpoint
        if (teamsFlow) {
          console.log("Redirecting to the login window.");
          userAgentApplication.loginRedirect({scopes: graphAPIScopes});
        } else {
          console.log("The user is not logged in and the redirect is disabled outside the authentication flow.");
        }
      } else {
        console.log("Getting token");
        this.getToken(userAgentApplication, graphAPIScopes, teamsFlow);
      }
    }
  }

  private getToken(userAgentApplication: Msal.UserAgentApplication, graphAPIScopes: string[], teamsFlow: boolean) {
    // In order to call the Microsoft Graph API, an access token needs to be acquired.
    // Try to acquire the token used to query Microsoft Graph API silently first:
    userAgentApplication.acquireTokenSilent({ scopes: graphAPIScopes }).then(
      (token) => {
        if (teamsFlow) {
          // After the access token is acquired, return to MS Teams, sending the acquired token
          microsoftTeams.authentication.notifySuccess(token.accessToken);
        } else {
          console.log("Silent token refresh successful.", token);
        }
      },
      (error) => {
        // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
        // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user
        // can reenter the current username/ password and/ or give consent to new permissions your application is requesting.
        if (error) {
          if (teamsFlow) {
            userAgentApplication.acquireTokenRedirect({ scopes: graphAPIScopes });
          } else {
            alert(error);
          }
        }
      }
      );
    }

    private tokenReceivedCallback(errorDesc, token, error, tokenType) {
      //  suppress typescript compile errors
    }
  }
