import { Injectable } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { Client } from '@microsoft/microsoft-graph-client';

import { AlertsService } from './alerts.service';
import { OAuthSettings } from '../oauth';
import { User } from './user';

@Injectable({
  providedIn: 'root'
})
export class AuthService {

  public authenticated: boolean;
  public user: User;

  constructor(
    private msalService: MsalService,
    private AlertsService: AlertsService ) {
    
    this.authenticated = this.msalService.getUser() !=null;
    this.getUser().then((user) => {this.user = user});
    }

async signIn(): Promise<void> {
  let result = await this.msalService.loginPopup(OAuthSettings.scopes)
  .catch((reason) => {
    this.AlertsService.add('Login Failed', JSON.stringify(reason,null,2));
  });
  
  if (result) {
    this.authenticated = true;
    this.user = await this.getUser();
  }

  }
  signOut(): void {
    this.msalService.logout();
    this.user = null;
    this.authenticated = false;
  }

  async getAccessToken(): Promise<string> {
    let result = await this.msalService.acquireTokenSilent(OAuthSettings.scopes)
      .catch((reason) => {
        this.AlertsService.add('Get token failed', JSON.stringify(reason, null, 2));
      });

    
    return result;
  }

  private async getUser(): Promise<User> {
    if (!this.authenticated) return null;

    let graphClient = Client.init({
      authProvider: async(done) => {
        let token = await this.getAccessToken()
        .catch((reason) => {
          done(reason, null);
        });

        if (token)
        {
          done(null,token);
        } else {
          done("Could not get access token", null);
        }
      }
    });

    let graphUser = await graphClient.api('/me').get();

    let user = new User();
    user.displayName = graphUser.displayName;
    user.email = graphUser.mail || graphUser.userPrincipalName;

    return user;
  }
}
