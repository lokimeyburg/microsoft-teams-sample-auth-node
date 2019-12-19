// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as constants from "./constants";
import * as storage from "./storage";
import * as utils from "./utils";
import { Request, Response } from "express";
import { RootDialog } from "./dialogs/RootDialog";
import { IOAuth2Provider } from "./providers";
import { logger } from "./utils";
import { UserToken } from "./providers";

// =========================================================
// Auth Bot
// =========================================================

export class AuthBot extends builder.UniversalBot {

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
        app: any,
    )
    {
        super(_connector, botSettings);
        this.set("persistConversationData", true);

        // Handle generic invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        teamsConnector.onInvoke(async (event, cb) => {
            try {
                await this.onInvoke(event, cb);
            } catch (e) {
                logger.error("Invoke handler failed", e);
                cb(e, null, 500);
            }
        });
        teamsConnector.onSigninStateVerification(async (event, query, cb) => {
            try {
                await this.onInvoke(event, cb);
            } catch (e) {
                logger.error("Signin state verification handler failed", e);
                cb(e, null, 500);
            }
        });

        // Register dialogs
        new RootDialog().register(this);
    }

    // Handle OAuth callbacks
    // The provider name is in the route, which is defined as "/auth/:provider/callback"
    public async handleOAuthCallback(req: Request, res: Response, providerName: string): Promise<void> {
        const provider = this.botSettings[providerName] as IOAuth2Provider;
        const stateString = req.query.state as string;
        const state = JSON.parse(stateString);
        const authCode = req.query.code;
        const accessToken = req.query.accessToken;
        let verificationCode = "";

        console.log("----------------------------")
        console.log(req.query)
        console.log("----------------------------")

        // Load the session from the address information in the OAuth state.
        // We'll later validate the state to check that it was not forged.
        let session: builder.Session;
        let address: builder.IAddress;
        try {
            address = state.address as builder.IAddress;
            session = await utils.loadSessionAsync(this, {
                type: "invoke",
                agent: "botbuilder",
                source: address.channelId,
                sourceEvent: {},
                address: address,
                user: address.user,
            });
        } catch (e) {
            logger.warn("Failed to get address from OAuth state", e);
        }

        console.log("------>")
        if ((utils.getOAuthState(session, providerName) === stateString)) {
            console.log("They match!")
        } else {
            console.log("They dont match")
        }
        console.log("State string:");
        console.log(stateString);
        console.log("OAuth State: ");
        console.log(utils.getOAuthState(session, providerName));
        console.log("Auth Code: ");
        console.log(authCode);
        console.log("<------")

        if (session &&
            (utils.getOAuthState(session, providerName) === stateString) &&     // OAuth state matches what we expect
            (authCode || accessToken)) {                                                         // User granted authorization
            try {
                // Redeem the authorization code for an access token (if we dont have one), 
                // and store it provisionally. The bot will refuse to use the token until we 
                // validate that the user in the chat is the same as the user who went through 
                // the authorization flow, using a verification code that needs to be presented 
                // by the user in the chat.
                if (accessToken != null) {
                    var userToken: UserToken = {
                        accessToken: accessToken,
                        expirationTime: 3957,
                        verificationCode: "asdasd"
                    }
                } else {
                    let userToken = await provider.getAccessTokenAsync(authCode);
                    await utils.prepareTokenForVerification(userToken);
                }

                utils.setUserToken(session, providerName, userToken);
                verificationCode = userToken.verificationCode;
            } catch (e) {
                logger.error("Failed to redeem code for an access token", e);
            }
        } else {
            logger.warn("State does not match expected state parameter, or user denied authorization");
        }

        // Render the page shown to the user
        if (verificationCode) {
            // If we have a verification code, we were able to redeem the code successfully. Render a page
            // that calls notifySuccess() with the verification code, or instructs the user to enter it in chat.
            res.render("oauth-callback-success", {
                verificationCode: verificationCode,
                providerName: provider.displayName,
            });

            // The auth flow resumes when we receive the verification code response, which can happen either:
            // 1) through notifySuccess(), which is handled in BaseIdentityDialog.handleLoginCallback()
            // 2) by user entering it in chat, which is handled in BaseIdentityDialog.onMessageReceived()

        } else {
            // Otherwise render an error page
            res.render("oauth-callback-error", {
                providerName: provider.displayName,
            });
        }
    }

    // Get the user's profile information from all the identity providers that we have tokens for
    public async getUserProfilesAsync(aadObjectId: string): Promise<any> {
        let profiles = {};

        let botStorage = this.get("storage") as storage.IBotExtendedStorage;
        let userData = await botStorage.getUserDataByAadObjectIdAsync(aadObjectId);
        if (userData) {
            for (let providerName in constants.IdentityProvider) {
                let token = utils.getUserTokenFromUserData(userData, providerName);
                let provider = this.get(providerName) as IOAuth2Provider;
                if (token && provider) {
                    let profile = await provider.getProfileAsync(token.accessToken);
                    profiles[provider.displayName] = profile;
                }
            }
        }

        return profiles;
    }

    // Handle incoming invoke
    private async onInvoke(event: builder.IEvent, cb: (err: Error, body: any, status?: number) => void): Promise<void> {
        let session = await utils.loadSessionAsync(this, event);
        if (session) {
            // Invokes don't participate in middleware

            // Simulate a normal message and route it, but remember the original invoke message
            let payload = (event as any).value;
            let fakeMessage: any = {
                ...event,
                text: payload.command + " " + JSON.stringify(payload),
                originalInvoke: event,
            };

            session.message = fakeMessage;
            session.dispatch(session.sessionState, session.message, () => {
                session.routeToActiveDialog();
            });
        }
        cb(null, "");
    }
}
