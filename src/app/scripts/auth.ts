import * as Msal from "msal";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * Implementation of the teams app1 Auth page
 */
export class Auth {
  private token: string = "";
  private user: Msal.Account;
  private graphAPIScopes: string[] = ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/User.ReadBasic.All", "https://graph.microsoft.com/Group.Read.All"];
  private uaa: Msal.UserAgentApplication;

  /**
   * Constructor for Tab that initializes the Microsoft Teams script
   */
  constructor() {
    microsoftTeams.initialize();

    microsoftTeams.getContext((context) => {
      const msalConfig: Msal.Configuration = {
        auth: {
          clientId: "e743c151-a549-4181-b3e9-e84052c9174c",
          authority: "https://login.microsoftonline.com/" + context.tid
        }
      };

      this.uaa = new Msal.UserAgentApplication(msalConfig);
      this.uaa.handleRedirectCallback(() => { const notUsed = ""; });
    });
  }

  public get onlineUser(): Msal.Account {
    return this.uaa.getAccount();
  }

  public performAuthV2(level: string) {
    if (this.uaa.isCallback(window.location.hash)) {
      const user = this.uaa.getAccount();
      if (user) {
        this.getToken(this.uaa, this.graphAPIScopes);
      }
    } else {
      this.user = this.uaa.getAccount();
      if (!this.user) {
        // If user is not signed in, then prompt user to sign in via loginRedirect.
        // This will redirect user to the Azure Active Directory v2 Endpoint
        this.uaa.loginRedirect({scopes: this.graphAPIScopes});
      } else {
        this.getToken(this.uaa, this.graphAPIScopes);
      }
    }
  }

  public getTheTokenAgain() {

    console.log("The User Agent Application", this.uaa);
    console.log("The User object", this.uaa.getAccount());

    console.log("Attempting to acquire token silently");
    this.uaa.acquireTokenSilent({ scopes: this.graphAPIScopes }).then(
      (token) => {
        // After the access token is acquired, return to MS Teams, sending the acquired token
        console.log("Success", token.accessToken);
      },
      (error) => {
        // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
        // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user
        // can reenter the current username/ password and/ or give consent to new permissions your application is requesting.
        if (error) {
          console.log("Error", error);
        }
      }
    );
  }

  private getToken(userAgentApplication: Msal.UserAgentApplication, graphAPIScopes: string[]) {
    // In order to call the Microsoft Graph API, an access token needs to be acquired.
    // Try to acquire the token used to query Microsoft Graph API silently first:
    userAgentApplication.acquireTokenSilent({ scopes: graphAPIScopes }).then(
      (token) => {
        // After the access token is acquired, return to MS Teams, sending the acquired token
        microsoftTeams.authentication.notifySuccess(token.accessToken);
      },
      (error) => {
        // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
        // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user
        // can reenter the current username/ password and/ or give consent to new permissions your application is requesting.
        if (error) {
          userAgentApplication.acquireTokenRedirect({ scopes: graphAPIScopes });
        }
      }
    );
  }

    private tokenReceivedCallback(errorDesc, token, error, tokenType) {
      console.log("a");
      //  suppress typescript compile errors
    }
  }
