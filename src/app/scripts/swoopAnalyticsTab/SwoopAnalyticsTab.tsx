import * as React from "react";
import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { Auth } from "../auth";

/**
 * State for the swoopAnalyticsTabTab React component
 */
export interface ISwoopAnalyticsTabState extends ITeamsBaseComponentState {
    entityId?: string;
    graphData?: string;
    teamsContext?: string;
    getTeamsUserResult?: any;
    tryToAuthenticateResult?: any;
}

/**
 * Properties for the swoopAnalyticsTabTab React component
 */
export interface ISwoopAnalyticsTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the SWOOP Analytics content page
 */
export class SwoopAnalyticsTab extends TeamsBaseComponent<ISwoopAnalyticsTabProps, ISwoopAnalyticsTabState> {

    private configuration?: string;
    private groupId?: string;
    private token?: string;

    public componentWillMount() {

        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.configuration = context.entityId;
                this.groupId = context.groupId;

                this.setState({
                    entityId: context.entityId,
                    teamsContext: JSON.stringify(context),
                    graphData: "Please click the button above to continue.",
                    getTeamsUserResult: "Please click the button above to continue.",
                });

                // this.tryToAuthenticate();
            });
        } else {
            this.setState({
                graphData: "This is not hosted in Microsoft Teams"
            });
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };

        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>SWOOP Analytics v10</div>
                        </PanelHeader>
                        <PanelBody>
                            <strong>Graph data</strong>
                            <div style={styles.section}>
                                <PrimaryButton onClick={() => this.getGraphData()}>Get Microsoft Graph data</PrimaryButton>
                            </div>
                            <div style={styles.section}>
                                {this.state.graphData}
                            </div>

                            <hr />
                            <PrimaryButton onClick={() => this.getTeamsUser()}>Get the current Teams user</PrimaryButton>
                            <div style={styles.section}>
                                {this.state.getTeamsUserResult}
                            </div>

                            <hr />
                            <PrimaryButton onClick={() => this.tryToAuthenticate()}>Try to get the token silently</PrimaryButton>
                            <div style={styles.section}>
                                {this.state.tryToAuthenticateResult}
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.footer}>
                                (C) Copyright SWOOP Analytics Pty Ltd
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }

    private getGraphData() {
        this.setState({
            graphData: "Loading..."
        });

        microsoftTeams.authentication.authenticate({
            url: "/auth.html",
            width: 400,
            height: 400,
            successCallback: (data) => {
                // Note: token is only good for one hour
                this.token = data!;
                this.getData(this.token);
            },
            failureCallback: (err) => {
                this.setState({
                    graphData: "Failed to authenticate and get token.<br/>" + err
                });
            }
        });
    }

    private tryToAuthenticate() {
        console.log("Performing authentication");
        const auth = new Auth();
        auth.performAuthV2(false);
    }

    private getTeamsUser() {
        microsoftTeams.authentication.getUser({
            successCallback: (user) => {
                this.setState({
                    getTeamsUserResult: JSON.stringify({
                        result: "getTeamsUser OK",
                        user
                    }, null, 2)
                });
            },
            failureCallback: (error) => {
                this.setState({
                    getTeamsUserResult: JSON.stringify({
                        result: "getTeamsUser FAIL",
                        error
                    }, null, 2)
                });
            }
        });
    }

    private getData(token: string) {
        let graphEndpoint = "https://graph.microsoft.com/v1.0/me";
        if (this.configuration === "GRP") {
            graphEndpoint = "https://graph.microsoft.com/v1.0/groups/" + this.groupId;
        }

        const req = new XMLHttpRequest();
        req.open("GET", graphEndpoint, false);
        req.setRequestHeader("Authorization", "Bearer " + token);
        req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");
        req.send();
        const result = JSON.parse(req.responseText);
        this.setState({
            graphData: JSON.stringify(result, null, 2)
        });
    }
}
