import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';

import { AuthService } from './auth.service';
import { Event } from './event';
import { AlertsService } from './alerts.service';
import {Contact} from './model/Contact';
import {ContactFolder} from './model/ContactFolder';
import {Message} from './model/Message';

@Injectable({
  providedIn: 'root'
})
export class GraphService {

  private graphClient: Client;
  constructor(
    private authService: AuthService,
    private alertsService: AlertsService) {

    // Initialize the Graph client
    this.graphClient = Client.init({
      authProvider: async (done) => {
        // Get the token from the auth service
        let token = await this.authService.getAccessToken()
          .catch((reason) => {
            done(reason, null);
          });

        if (token)
        {
          done(null, token);
        } else {
          done("Could not get an access token", null);
        }
      }
    });
  }

  async getEvents(): Promise<Event[]> {
    try {
      let result =  await this.graphClient
        .api('/me/events')
        .select('subject,organizer,start,end')
        .orderby('createdDateTime DESC')
        .get();

      return result.value;
    } catch (error) {
      this.alertsService.add('Could not get events', JSON.stringify(error, null, 2));
    }
  }

  async getAll(): Promise<ContactFolder[]> {
    try {
      let result = await this.graphClient
        .api('me/contactFolders')
        .get();
      return result.value;
    } catch(error) {
      this.alertsService.add('Could not get subFolders id', JSON.stringify(error, null, 2));
    }
  }

  async getContacts(id: string): Promise<Contact[]> {
    try {
       let result = await this.graphClient
          .api('me/contactFolders/' + id + '/contacts')
          .select('displayName,emailAddresses')
          .get();
      return result.value;
    } catch (error) {
      this.alertsService.add('Could not get List of Contacts', JSON.stringify(error, null, 2));
    }
  }

  async getDefaultContacts(): Promise<Contact[]> {
    try {
      let result = await this.graphClient
        .api('me/contacts')
        .select('displayName,emailAddresses')
        .get();
      return result.value;
    } catch (error) {
      this.alertsService.add('Could not get default contacts', JSON.stringify(error, null, 2));
    }
  }

  async sendMail(mail: Message): Promise<any> {
    await this.graphClient
      .api('me/sendMail')
      .post(JSON.stringify(mail));
  } catch(error) {
    this.alertsService.add('Could not send mail', JSON.stringify(error, null, 2));
  }
}
