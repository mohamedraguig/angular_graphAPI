import { Component, OnInit } from '@angular/core';
import {Contact} from '../model/Contact';
import {GraphService} from '../graph.service';
import {ContactFolder} from '../model/ContactFolder';
import {FormControl, Validators} from '@angular/forms';
import {Body, Message, MessageClass, ToRecipient} from '../model/Message';
import {EmailAddress} from '../model/EmailAddress';

@Component({
  selector: 'app-contacts',
  templateUrl: './contacts.component.html',
  styleUrls: ['./contacts.component.css']
})
export class ContactsComponent implements OnInit {

  private contacts: Contact[] = [];
  private contactFolders: ContactFolder[] = [];
  private email = new FormControl('', [Validators.email]);
  contactsList: string[];

  constructor(private graphService: GraphService) { }

  ngOnInit() {
    this.graphService.getAll().then((result) => {
      this.contactFolders = result;

      if (this.contactFolders.length > 0) {
        for ( let i = 0; i < this.contactFolders.length; i++) {
          this.graphService.getContacts(this.contactFolders[i].id).then((contactResult) => {
            contactResult.map(val => this.contacts.push(Object.assign({}, val)));
          });
        }
        this.getDefaultContacts();
      } else {
        this.getDefaultContacts();
      }
    });
  }

  getDefaultContacts() {
    this.graphService.getDefaultContacts().then((defaultContacts) => {
      defaultContacts.map(val => this.contacts.push(Object.assign({}, val)));
    });
  }

  getErrorMessage() {
    return this.email.hasError('email') ? 'Email non valide' : '';
  }

  sendMail(destinataires: string, objet: string, message: string) {
    console.log('Resultats',destinataires+objet+message);

    let messageToSend: Message = new Message();
    let mailToSend: MessageClass = new MessageClass();
    let mailBody: Body = new Body();
    let finalRecipients: ToRecipient[] = [];

    messageToSend.saveToSentItems = 'true';

    mailBody.contentType = 'Text';
    mailBody.content = message;
    mailToSend.body = mailBody;
    mailToSend.subject = objet;

    destinataires.split(',').map(mailAddress => {
      let emailAddress: EmailAddress = new EmailAddress();
      let recipients: ToRecipient = new ToRecipient();
      emailAddress.address = mailAddress;
      emailAddress.name = mailAddress;
      recipients.emailAddress = emailAddress;
      finalRecipients.push(recipients);
    });
    mailToSend.toRecipients = finalRecipients;

    console.log('Liste des destinataires', mailToSend);

    messageToSend.message = mailToSend;

    this.graphService.sendMail(messageToSend).then((resultat) => {
      console.log('Message sent');
    });
  }

}
