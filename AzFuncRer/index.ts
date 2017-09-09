import { Parser } from "xml2js";
import { TokenHelper, IOnlineAddinCredentials, IAuthData } from "node-sp-auth";
import * as pnp from "sp-pnp-js";
import * as nodeFetch from 'node-fetch';
import fetch from 'node-fetch';
import * as fs from 'fs';

declare var global: any;

export function run(context: any, req: any): void {
    context.log("Running Remote Event Receiver from Azure Function!");

    configurePnP();

    execute(context, req)
        .catch((err: any) => {
            console.log(err);
            context.done();
        });
}

async function execute(context: any, req: any): Promise<any> {
    let data = await xml2Json(req.body);

    if(data["s:Envelope"]["s:Body"].ProcessOneWayEvent){
        await processOneWayEvent(data["s:Envelope"]["s:Body"].ProcessOneWayEvent.properties, context);
    } else if(data["s:Envelope"]["s:Body"].ProcessEvent){
        await processEvent(data["s:Envelope"]["s:Body"].ProcessEvent.properties, context)
    } else {
        throw new Error("Unable to resolve event type");
    }
}

async function processEvent(eventProperties: any, context: any): Promise<any> {
    // for demo: cancel sync -ing RER with error:

    let body = fs.readFileSync('AzFuncRer/response.data').toString();
    context.res = {
        status: 200,
        headers: {
            "Content-Type": "text/xml"
        },
        body: body,
        isRaw: true
    } as any;

    context.done();
}

async function processOneWayEvent(eventProperties: any, context: any): Promise<any> {
    let contextToken = eventProperties.ContextToken;
    let itemProperties = eventProperties.ItemEventProperties;

    let creds: IOnlineAddinCredentials = {
        clientId: getAppSetting('ClientId'),
        clientSecret: getAppSetting('ClientSecret')
    };

    let appToken = TokenHelper.verifyAppToken(contextToken, creds);
    let authData: IAuthData = {
        refreshToken: appToken.refreshtoken,
        realm: appToken.realm,
        securityTokenServiceUri: appToken.context.SecurityTokenServiceUri
    };

    let acessToken = await TokenHelper.getUserAccessToken(itemProperties.WebUrl, authData, creds);
    let sp = pnp.sp.configure({
        headers: {
            "Authorization": `Bearer ${acessToken.value}`
        }
    }, itemProperties.WebUrl);

    let itemUpdate = await sp.web.lists.getById(itemProperties.ListId).items.getById(itemProperties.ListItemId)
        .update({
            Title: "Updated by Azure function!"
        });

    console.log(itemUpdate);

    context.res = {
        status: 200,
        body: ''
    } as any;

    context.done();
}

async function xml2Json(input: string): Promise<any> {
    return new Promise((resolve, reject) => {
        let parser = new Parser({
            explicitArray: false
        });

        parser.parseString(input, (jsError: any, jsResult: any) => {
            if (jsError) {
                reject(jsError);
            } else {
                resolve(jsResult);
            }
        });
    });
}

function configurePnP(): void {
    global.Headers = nodeFetch.Headers;
    global.Request = nodeFetch.Request;
    global.Response = nodeFetch.Response;

    pnp.setup({
        fetchClientFactory: () => {
            return {
                fetch: (url: string, options: any): Promise<any> => {
                    return fetch(url, options);
                }
            }
        }
    })
}

function getAppSetting(name: string): string {
    return process.env[name] as string;
}