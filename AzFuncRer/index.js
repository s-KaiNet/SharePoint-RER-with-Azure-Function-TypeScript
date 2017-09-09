"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const xml2js_1 = require("xml2js");
const node_sp_auth_1 = require("node-sp-auth");
const pnp = require("sp-pnp-js");
const nodeFetch = require("node-fetch");
const node_fetch_1 = require("node-fetch");
const fs = require("fs");
function run(context, req) {
    context.log("Running Remote Event Receiver from Azure Function!");
    configurePnP();
    execute(context, req)
        .catch((err) => {
        console.log(err);
        context.done();
    });
}
exports.run = run;
function execute(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        let data = yield xml2Json(req.body);
        if (data["s:Envelope"]["s:Body"].ProcessOneWayEvent) {
            yield processOneWayEvent(data["s:Envelope"]["s:Body"].ProcessOneWayEvent.properties, context);
        }
        else if (data["s:Envelope"]["s:Body"].ProcessEvent) {
            yield processEvent(data["s:Envelope"]["s:Body"].ProcessEvent.properties, context);
        }
        else {
            throw new Error("Unable to resolve event type");
        }
    });
}
function processEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        // for demo: cancel sync -ing RER with error:
        let body = fs.readFileSync('AzFuncRer/response.data').toString();
        context.res = {
            status: 200,
            headers: {
                "Content-Type": "text/xml"
            },
            body: body,
            isRaw: true
        };
        context.done();
    });
}
function processOneWayEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        let contextToken = eventProperties.ContextToken;
        let itemProperties = eventProperties.ItemEventProperties;
        let creds = {
            clientId: getAppSetting('ClientId'),
            clientSecret: getAppSetting('ClientSecret')
        };
        let appToken = node_sp_auth_1.TokenHelper.verifyAppToken(contextToken, creds);
        let authData = {
            refreshToken: appToken.refreshtoken,
            realm: appToken.realm,
            securityTokenServiceUri: appToken.context.SecurityTokenServiceUri
        };
        let acessToken = yield node_sp_auth_1.TokenHelper.getUserAccessToken(itemProperties.WebUrl, authData, creds);
        let sp = pnp.sp.configure({
            headers: {
                "Authorization": `Bearer ${acessToken.value}`
            }
        }, itemProperties.WebUrl);
        let itemUpdate = yield sp.web.lists.getById(itemProperties.ListId).items.getById(itemProperties.ListItemId)
            .update({
            Title: "Updated by Azure function!"
        });
        console.log(itemUpdate);
        context.res = {
            status: 200,
            body: ''
        };
        context.done();
    });
}
function xml2Json(input) {
    return __awaiter(this, void 0, void 0, function* () {
        return new Promise((resolve, reject) => {
            let parser = new xml2js_1.Parser({
                explicitArray: false
            });
            parser.parseString(input, (jsError, jsResult) => {
                if (jsError) {
                    reject(jsError);
                }
                else {
                    resolve(jsResult);
                }
            });
        });
    });
}
function configurePnP() {
    global.Headers = nodeFetch.Headers;
    global.Request = nodeFetch.Request;
    global.Response = nodeFetch.Response;
    pnp.setup({
        fetchClientFactory: () => {
            return {
                fetch: (url, options) => {
                    return node_fetch_1.default(url, options);
                }
            };
        }
    });
}
function getAppSetting(name) {
    return process.env[name];
}
//# sourceMappingURL=index.js.map