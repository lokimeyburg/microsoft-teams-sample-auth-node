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

let express = require("express");
let exphbs  = require("express-handlebars");
import { Request, Response } from "express";
let bodyParser = require("body-parser");
let favicon = require("serve-favicon");
let http = require("http");
let path = require("path");
let querystring = require("querystring");
let fetch = require("node-fetch");
import * as config from "config";
import * as msteams from "botbuilder-teams";
import * as apis from "./apis";
import * as providers from "./providers";
import * as storage from "./storage";
import { AuthBot } from "./AuthBot";
import { logger } from "./utils/index";

let app = express();
let appId = config.get("app.appId");

app.set("port", process.env.PORT || 3333);
app.use(express.static(path.join(__dirname, "../../public")));
app.use(favicon(path.join(__dirname, "../../public/assets", "favicon.ico")));
app.use(bodyParser.json());

let handlebars = exphbs.create({
    extname: ".hbs",
    helpers: {
        appId: () => { return appId; },
    },
    defaultLayout: false,
});
app.engine("hbs", handlebars.engine);
app.set("view engine", "hbs");

// Configure storage
let botStorageProvider = config.get("storage");
let botStorage = null;
switch (botStorageProvider) {
    case "mongoDb":
        botStorage = new storage.MongoDbBotStorage(config.get("mongoDb.botStateCollection"), config.get("mongoDb.connectionString"));
        break;
    case "memory":
        botStorage = new storage.MemoryBotStorage();
        break;
    case "null":
        botStorage = new storage.NullBotStorage();
        break;
}

// Create chat bot
let connector = new msteams.TeamsChatConnector({
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
});
let botSettings = {
    storage: botStorage,
    linkedIn: new providers.LinkedInProvider(config.get("linkedIn.clientId"), config.get("linkedIn.clientSecret")),
    azureADv1: new providers.AzureADv1Provider(config.get("azureAD.appId"), config.get("azureAD.appPassword")),
    google: new providers.GoogleProvider(config.get("google.clientId"), config.get("google.clientSecret")),
};
let bot = new AuthBot(connector, botSettings, app);

// Log bot errors
bot.on("error", (error: Error) => {
    logger.error(error.message, error);
});

// Configure bot routes
app.post("/api/messages", connector.listen());

// Configure auth callback routes
app.get("/auth/:provider/callback", (req, res) => {
    bot.handleOAuthCallback(req, res, req.params["provider"]);
});

// Tab authentication sample routes
app.get("/tab/simple", (req, res) => { res.render("tab/simple/simple"); });
app.get("/tab/simple-start", (req, res) => { res.render("tab/simple/simple-start"); });
app.get("/tab/simple-start-v2", (req, res) => { res.render("tab/simple/simple-start-v2"); });
app.get("/tab/simple-end", (req, res) => { res.render("tab/simple/simple-end"); });
app.get("/tab/silent", (req, res) => { res.render("tab/silent/silent"); });
app.get("/tab/silent-start", (req, res) => { res.render("tab/silent/silent-start"); });
app.get("/tab/silent-end", (req, res) => { res.render("tab/silent/silent-end"); });

// On-behalf-of token exchange
app.post("/auth/token", (req, res) => {
    let tid = req.body.tid;
    let token = req.body.token;
    let scopes = ["https://graph.microsoft.com/User.Read"];
    let oboPromise = new Promise((resolve, reject) => {
        const url = "https://login.microsoftonline.com/" + tid + "/oauth2/v2.0/token";
        const params = {
            client_id: "bdb71ee3-1c28-4edb-a758-fd6f8b60348c",
            client_secret: "]DjvGB0f?R[Z4qSwn24uSfr?EKhGN_tv",
            grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
            assertion: token,
            requested_token_use: "on_behalf_of",
            scope: scopes.join(" "),
        };

        fetch(url, {
            method: "POST",
            body: querystring.stringify(params),
            headers: {
                Accept: "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
            },
        }).then(result => {
            if (result.status !== 200) {
            result.json().then(json => {
                // TODO: Check explicitly for invalid_grant or interaction_required
                reject({"error": json.error});
            });
            } else {
            result.json().then(json => {
                resolve(json);
            });
            }
        });
    });

    oboPromise.then(result => {
        res.json(result);
    }, err => {
        console.log(err); // Error: "It broke"
        res.json(err);
    });
});

let openIdMetadata = new apis.OpenIdMetadata("https://login.microsoftonline.com/common/.well-known/openid-configuration");
let validateIdToken = new apis.ValidateIdToken(openIdMetadata, appId).listen();     // Middleware to validate id_token
app.get("/api/decodeToken", validateIdToken, new apis.DecodeIdToken().listen());
app.get("/api/getProfileFromGraph", validateIdToken, new apis.GetProfileFromGraph(config.get("app.appId"), config.get("app.appPassword")).listen());
app.get("/api/getProfilesFromBot", validateIdToken, async (req, res) => {
    let profiles = await bot.getUserProfilesAsync(res.locals.token["oid"]);
    res.status(200).send(profiles);
});

// Configure ping route
app.get("/ping", (req, res) => {
    res.status(200).send("OK");
});

// error handlers

// development error handler
// will print stacktrace
if (app.get("env") === "development") {
    app.use(function(err: any, req: Request, res: Response, next: Function): void {
        logger.error("Failed request", err);
        res.status(err.status || 500).send(err);
    });
}

// production error handler
// no stacktraces leaked to user
app.use(function(err: any, req: Request, res: Response, next: Function): void {
    logger.error("Failed request", err);
    res.sendStatus(err.status || 500);
});

http.createServer(app).listen(app.get("port"), function (): void {
    logger.verbose("Express server listening on port " + app.get("port"));
    logger.verbose("Bot messaging endpoint: " + config.get("app.baseUri") + "/api/messages");
});
