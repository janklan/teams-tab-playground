import * as Msal from "msal";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * Implementation of the teams app1 Auth page
 */
export class Auth {
  private token: string = "";
  private user: Msal.Account;

  private msalApp;

  /**
   * Constructor for Tab that initializes the Microsoft Teams script
   */
  constructor() {
    microsoftTeams.initialize();

    const msalConfig: Msal.Configuration = {
      auth: {
        clientId: "e743c151-a549-4181-b3e9-e84052c9174c",
        authority: "https://login.microsoftonline.com/c5870e0f-a946-4008-9f5c-94875cba8b2e" // todo replace with teams context tid
      }
    };

    this.msalApp = new Msal.UserAgentApplication(msalConfig);
    this.msalApp.handleRedirectCallback(() => { const notUsed = ""; });
  }

  public performAuthV2MSALPopup() {

    const tokenRequest = {
      scopes: ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/User.ReadBasic.All", "email", "profile", "openid"]
    };

    // if the user is already logged in you can acquire a token
    if (this.msalApp.getAccount()) {
      this.msalApp.acquireTokenSilent(tokenRequest)
        .then(response => {
          console.log("Response0", response);
          // get access token from response
          // response.accessToken
        })
        .catch(err => {
          // could also check if err instance of InteractionRequiredAuthError if you can import the class.
          if (err.name === "InteractionRequiredAuthError") {
            return this.msalApp.acquireTokenPopup(tokenRequest)
              .then(response => {
                console.log("Response1", response);
                // get access token from response
                // response.accessToken
              })
              .catch(err => {
                console.log("Error1", err);
                // handle error
              });
          }
        });
    } else {
      return this.msalApp.acquireTokenPopup(tokenRequest)
              .then(response => {
                console.log("Response2", response);
                // get access token from response
                // response.accessToken
              })
              .catch(err => {
                console.log("Error2", err);
                // handle error
              });
    }
  }

  public performAuthV2(teamsFlow: boolean = true) {
    console.log("Authv2", teamsFlow);
    // Setup auth parameters for MSAL
    const graphAPIScopes: string[] = ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/User.ReadBasic.All", "email", "profile", "openid"];

    const urlParams = new URLSearchParams(location.search);
    console.log(urlParams);

    if (!this.msalApp) {
      alert("The MSAL app isn't ready yet");
    }

    const userAgentApplication = this.msalApp; // ugly, but done to avoid changing lots of the stock code

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
          console.log("The user is not logged. If this was a popup, the user would be redirected.");
          console.log("Getting token anyway.");
          this.getToken(userAgentApplication, graphAPIScopes, teamsFlow);
        }
      } else {
        console.log("Getting token");
        this.getToken(userAgentApplication, graphAPIScopes, teamsFlow);
      }
    }
  }

  private async getTokenSilent(userAgentApplication: Msal.UserAgentApplication, graphAPIScopes: string[], teamsFlow: boolean) {
    console.log("getTokenSilent: Begin");

    try {
      const token = await userAgentApplication.acquireTokenSilent({ scopes: graphAPIScopes });
      console.log("getTokenSilent: Token acquired", token);
    } catch (error) {
      console.log("getTokenSilent: Error getting the token silently", error);
    }

    console.log("getTokenSilent: End");
  }

  private getToken(userAgentApplication: Msal.UserAgentApplication, graphAPIScopes: string[], teamsFlow: boolean) {
    console.log("getToken: Begin");

    // In order to call the Microsoft Graph API, an access token needs to be acquired.
    // Try to acquire the token used to query Microsoft Graph API silently first:
    userAgentApplication.acquireTokenSilent({ scopes: graphAPIScopes }).then(
      (token) => {
        console.log("getToken: Token acquired", token);
        if (teamsFlow) {
          // After the access token is acquired, return to MS Teams, sending the acquired token
          microsoftTeams.authentication.notifySuccess(token.accessToken);
        }
      },
      (error) => {
        console.log("getToken: Error getting the token silently", error);
        // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
        // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user
        // can reenter the current username/ password and/ or give consent to new permissions your application is requesting.
        if (teamsFlow) {
          userAgentApplication.acquireTokenRedirect({ scopes: graphAPIScopes });
        }
      }
    );

    console.log("getToken: End (but that doesn't mean the MSAL library is done)");
  }

  private tokenReceivedCallback(errorDesc, token, error, tokenType) {
    //  suppress typescript compile errors
  }
}
