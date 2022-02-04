import { HttpHeaders, HttpParams } from "@angular/common/http";
import { Injectable } from "@angular/core";
import { AuthenticationService } from "@sinequa/core/login";

import * as microsoftTeams from "@microsoft/teams-js";

@Injectable({providedIn: "root"})
export class TeamsAuthenticationService extends AuthenticationService {

    teamsToken: string;

    addAuthentication(config: {headers: HttpHeaders, params: HttpParams}) {
        config = super.addAuthentication(config);
        if(this.teamsToken) {
            config.headers = config.headers.set("teams-token", this.teamsToken);
        }
        return config;
    }

}

export function TeamsInitializer(authService: TeamsAuthenticationService): () => Promise<boolean> {
    
    const init = () => new Promise<boolean>((resolve, reject) => {
        if(!inIframe()) {
            resolve(true);
        }
       
        const authTokenRequest = {
            successCallback: result => {
                console.log("success callback:", result);
                if(authService) {
                    authService.teamsToken = result;
                }
                resolve(true);
            },
            failureCallback: error => {
                console.error("failure callback:", error);
                reject(error);
            }
        }

        console.log("Teams init", microsoftTeams);
        microsoftTeams.initialize(() => {
            console.log("teams initialized");
            microsoftTeams.authentication.getAuthToken(authTokenRequest);
        });
        microsoftTeams.getContext(context => console.log("Context",context));
    });

    return init;
}

export function inIframe() {
    try {
        return window.self !== window.top;
    } catch (e) {
        return true;
    }
}